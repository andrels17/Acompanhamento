import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import textwrap
import io
import os

# ---------------- Configurações ----------------
EXCEL_PATH = "Acompto_Abast.xlsx"

# Paletas (clara e escura)
PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# Classes que serão agrupadas em "Outros"
OUTROS_CLASSES = {"Motocicletas", "Mini Carregadeira", "Usina", "Veiculos Leves"}

# NOVA CONFIGURAÇÃO: Intervalos por tipo de equipamento
INTERVALOS_TIPO = {
    'HORAS': {  # Para máquinas agrícolas
        'lubrificacao': 250,
        'revisao': 1000
    },
    'QUILÔMETROS': {  # Para caminhões e veículos
        'lubrificacao': 5000,
        'revisao': 10000
    }
}

# Thresholds de alerta (quando avisar que está próximo)
ALERTAS_TIPO = {
    'HORAS': {
        'lubrificacao': 20,  # avisar quando faltarem 20h
        'revisao': 50        # avisar quando faltarem 50h
    },
    'QUILÔMETROS': {
        'lubrificacao': 500,  # avisar quando faltarem 500km
        'revisao': 1000      # avisar quando faltarem 1000km
    }
}

# ---------------- Utilitários ----------------
def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def wrap_labels(s: str, width: int = 18) -> str:
    """Quebra um rótulo em múltiplas linhas usando <br> para Plotly."""
    if pd.isna(s):
        return ""
    parts = textwrap.wrap(str(s), width=width)
    return "<br>".join(parts) if parts else str(s)

# NOVA FUNÇÃO: Detecta tipo de equipamento baseado na coluna Unid
def detect_equipment_type(df_abast: pd.DataFrame) -> pd.DataFrame:
    """Detecta se equipamento é controlado por HORAS ou QUILÔMETROS baseado na coluna Unid."""
    df = df_abast.copy()
    
    # Mapeia a coluna Unid para tipo de controle
    df['Tipo_Controle'] = df['Unid'].map({
        'HORAS': 'HORAS',
        'QUILÔMETROS': 'QUILÔMETROS'
    })
    
    # Para equipamentos sem informação na coluna Unid, tenta inferir pela classe
    # (você pode ajustar essas regras conforme sua necessidade)
    def inferir_tipo_por_classe(row):
        if pd.notna(row['Tipo_Controle']):
            return row['Tipo_Controle']
        
        classe = str(row.get('Classe_Operacional', '')).upper()
        
        # Máquinas agrícolas geralmente são por HORAS
        if any(palavra in classe for palavra in ['TRATOR', 'COLHEITADEIRA', 'PULVERIZADOR', 'PLANTADEIRA']):
            return 'HORAS'
        
        # Veículos geralmente são por QUILÔMETROS  
        if any(palavra in classe for palavra in ['CAMINHÃO', 'CAMINHAO', 'VEICULO', 'PICKUP']):
            return 'QUILÔMETROS'
        
        # Default para HORAS (máquinas agrícolas são maioria)
        return 'HORAS'
    
    df['Tipo_Controle'] = df.apply(inferir_tipo_por_classe, axis=1)
    
    return df

# Leitura segura do Excel (usa pandas). Cache para performance
@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Carrega e prepara DataFrames (Abastecimento e Frotas)."""
    try:
        df_abast = pd.read_excel(path, sheet_name="BD", skiprows=2)
        df_frotas = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()
    except ValueError as e:
        if "Sheet name" in str(e):
            st.error("Verifique se as planilhas 'BD' e 'FROTAS' existem no arquivo.")
            st.stop()
        else:
            raise

    # Normaliza frotas
    df_frotas = df_frotas.rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"}).drop_duplicates(subset=["Cod_Equip"])
    df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas.get("ANOMODELO", pd.Series()), errors="coerce")
    df_frotas["label"] = (
        df_frotas["Cod_Equip"].astype(str)
        + " - "
        + df_frotas.get("DESCRICAO_EQUIPAMENTO", "").fillna("")
        + " ("
        + df_frotas.get("PLACA", "").fillna("Sem Placa")
        + ")"
    )

    # --- CORREÇÃO APLICADA AQUI ---
    # A lista de nomes agora tem 21 colunas, exatamente como na sua planilha.
    # A coluna "Classe_Original" foi removida.
    df_abast.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes", "Semana",
        "Classe_Operacional", "Descricao_Proprietario_Original",
        "Potencia_CV_Abast", "Hod_Hor_Atual", "Unid"
    ]

    # APLICA A DETECÇÃO DE TIPO DE EQUIPAMENTO
    df_abast = detect_equipment_type(df_abast)

    df = pd.merge(df_abast, df_frotas, on="Cod_Equip", how="left")
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df.dropna(subset=["Data"], inplace=True)

    # Campos de tempo/derivados
    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Numéricos
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod", "Hod_Hor_Atual"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Marca / Fazenda (mantém coluna, mas não será usada em filtros)
    df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    df["Fazenda"] = df["Ref1"].astype(str)

    # Cálculo seguro de Consumo km/l (fallback)
    if "Km_Hs_Rod" in df.columns and "Qtde_Litros" in df.columns:
        df["Consumo_km_l"] = np.where(df["Qtde_Litros"] > 0, df["Km_Hs_Rod"] / df["Qtde_Litros"], np.nan)
        df["Media"] = df["Consumo_km_l"]
    else:
        df["Consumo_km_l"] = np.nan

    return df, df_frotas
@st.cache_data
def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Filtra o DataFrame conforme opções selecionadas (sem filtro de marca)."""
    mask = (
        df["Safra"].isin(opts["safras"]) &
        df["Ano"].isin(opts["anos"]) &
        df["Mes"].isin(opts["meses"]) &
        df["Classe_Operacional"].isin(opts["classes_op"])
    )
    return df.loc[mask].copy()

@st.cache_data
def calcular_kpis_consumo(df: pd.DataFrame) -> dict:
    """Calcula KPIs principais (total, média, equipamentos únicos)."""
    return {
        "total_litros": float(df["Qtde_Litros"].sum()) if "Qtde_Litros" in df.columns else 0.0,
        "media_consumo": float(df["Media"].mean()) if "Media" in df.columns else 0.0,
        "eqp_unicos": int(df["Cod_Equip"].nunique()) if "Cod_Equip" in df.columns else 0,
    }

def make_bar(fig_df, x, y, title, labels, palette, rotate_x=-60, ticksize=10, height=None, hoverfmt=None, wrap_width=18, hide_text_if_gt=8):
    """Helper para criar barras padronizadas com hovertemplate e rótulos de X legíveis."""
    df_local = fig_df.copy()
    if x in df_local.columns:
        df_local[x] = df_local[x].astype(str).apply(lambda s: wrap_labels(s, width=wrap_width))

    fig = px.bar(df_local, x=x, y=y, text=y, title=title, labels=labels, color_discrete_sequence=palette)
    # decide mostrar texto nas barras dependendo do número de categorias
    if df_local.shape[0] > hide_text_if_gt:
        fig.update_traces(texttemplate=None)
    else:
        fig.update_traces(texttemplate="%{text:.1f}", textfont=dict(size=10))

    fig.update_layout(
        xaxis=dict(tickangle=rotate_x, tickfont=dict(size=ticksize), automargin=True),
        margin=dict(l=40, r=20, t=60, b=160),
        title=dict(x=0.01, xanchor="left"),
        font=dict(size=13)
    )
    if height:
        fig.update_layout(height=height)
    if hoverfmt:
        fig.update_traces(hovertemplate=hoverfmt)
    else:
        fig.update_traces(hovertemplate=None)
    return fig

# ---------------- Layout / CSS moderno ----------------
def apply_modern_css(dark: bool):
    """Aplica CSS leve para um visual mais moderno."""
    bg = "#0e1117" if dark else "#FFFFFF"
    card_bg = "#111318" if dark else "#f8f9fa"
    text_color = "#f0f0f0" if dark else "#111111"
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-color: {bg};
            color: {text_color};
        }}
        .kpi-card {{
            background: {card_bg};
            padding: 12px;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        }}
        .kpi-title {{ font-size:14px; color: {text_color}; opacity:0.9 }}
        .kpi-value {{ font-size:20px; font-weight:700; color: {text_color} }}
        .section-title {{ font-size:18px; font-weight:700; color: {text_color} }}
        .small-muted {{ color: #8a8a8a; font-size:12px; }}
        .maintenance-alert {{ background: #fee2e2; border: 1px solid #fecaca; padding: 8px; border-radius: 6px; }}
        .maintenance-warning {{ background: #fef3c7; border: 1px solid #fde68a; padding: 8px; border-radius: 6px; }}
        </style>
        """,
        unsafe_allow_html=True
    )

# ---------------- NOVA Manutenção: lógica atualizada ----------------
def get_current_value_from_bd(df_abast: pd.DataFrame) -> pd.Series:
    """Pega o valor atual (mais recente) da coluna Hod_Hor_Atual por equipamento."""
    # Agrupa por equipamento e pega o último valor de Hod_Hor_Atual
    current_values = (
        df_abast
        .sort_values(["Cod_Equip", "Data"])
        .groupby("Cod_Equip")
        .agg({
            "Hod_Hor_Atual": "last",
            "Tipo_Controle": "last",
            "Classe_Operacional": "last"
        })
    )
    return current_values

def build_maintenance_table_new(df_abast: pd.DataFrame, df_frotas: pd.DataFrame) -> pd.DataFrame:
    """Constrói tabela com próximos serviços baseada no tipo de equipamento e coluna Hod_Hor_Atual."""
    
    # Pega valores atuais da coluna Hod_Hor_Atual
    current_data = get_current_value_from_bd(df_abast)
    
    # Merge com frotas para ter informações completas
    mf = df_frotas.copy()
    mf = mf.merge(current_data, left_on="Cod_Equip", right_index=True, how="left")
    
    # Função para calcular próximos serviços
    def calcular_proximos_servicos(row):
        tipo_controle = row.get("Tipo_Controle") # Não usar default aqui para pegar NaN
        valor_atual = row.get("Hod_Hor_Atual", np.nan)
        
        # --- CORREÇÃO APLICADA AQUI ---
        # A condição de entrada do IF agora lida com tipo_controle sendo NaN
        if pd.isna(valor_atual) or pd.isna(tipo_controle) or tipo_controle not in INTERVALOS_TIPO:
            return {
                "Prox_Lubrificacao": np.nan,
                "Prox_Revisao": np.nan,
                "Para_Lubrificacao": np.nan,
                "Para_Revisao": np.nan,
                "Alert_Lubrificacao": False,
                "Alert_Revisao": False,
                # Verifica se tipo_controle é texto antes de usar .lower()
                "Unidade": tipo_controle.lower() if isinstance(tipo_controle, str) else ""
            }
        
        # Pega intervalos para este tipo
        intervalos = INTERVALOS_TIPO[tipo_controle]
        alertas = ALERTAS_TIPO[tipo_controle]
        
        # Calcula próximos valores
        prox_lub = ((valor_atual // intervalos["lubrificacao"]) + 1) * intervalos["lubrificacao"]
        prox_rev = ((valor_atual // intervalos["revisao"]) + 1) * intervalos["revisao"]
        
        para_lub = prox_lub - valor_atual
        para_rev = prox_rev - valor_atual
        
        # Verifica se precisa de alerta
        alert_lub = para_lub <= alertas["lubrificacao"]
        alert_rev = para_rev <= alertas["revisao"]
        
        unidade = "h" if tipo_controle == "HORAS" else "km"
        
        return {
            "Prox_Lubrificacao": prox_lub,
            "Prox_Revisao": prox_rev,
            "Para_Lubrificacao": para_lub,
            "Para_Revisao": para_rev,
            "Alert_Lubrificacao": alert_lub,
            "Alert_Revisao": alert_rev,
            "Unidade": unidade,
            "Intervalo_Lub": intervalos["lubrificacao"],
            "Intervalo_Rev": intervalos["revisao"]
        }
    
    # Aplica cálculos
    calc_results = mf.apply(calcular_proximos_servicos, axis=1, result_type='expand')
    mf = pd.concat([mf.drop(columns=calc_results.columns, errors='ignore'), calc_results], axis=1)
    
    # Flag geral de alerta
    mf["Qualquer_Alerta"] = mf["Alert_Lubrificacao"] | mf["Alert_Revisao"]
    
    return mf

# ---------------- Excel I/O: salvar log e atualizar planilha ----------------
def read_all_sheets(path: str) -> dict:
    """Lê todas as abas do Excel em um dict {sheetname: dataframe}."""
    if not os.path.exists(path):
        return {}
    try:
        all_sheets = pd.read_excel(path, sheet_name=None)
        return all_sheets
    except Exception as e:
        st.error(f"Erro ao ler o Excel: {e}")
        return {}

def save_all_sheets(path: str, sheets: dict):
    """Sobrescreve o arquivo Excel com o dict de sheets."""
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
            for name, df in sheets.items():
                # evitar índices desnecessários
                df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar o Excel: {e}")
        raise

def append_manut_log_new(path: str, action: dict):
    """
    action: dict com keys:
    - Cod_Equip, DESCRICAO_EQUIPAMENTO, Tipo_Servico (Lubrificacao/Revisao), Valor_Atual, Tipo_Controle, Observacao, Data
    """
    sheets = read_all_sheets(path)
    if sheets is None:
        sheets = {}
    log_df = None
    if "MANUT_LOG" in sheets:
        log_df = sheets["MANUT_LOG"]
    else:
        # cria colunas básicas
        cols = ["Data", "Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Servico", "Valor_Atual", "Tipo_Controle", "Observacao", "Usuario"]
        log_df = pd.DataFrame(columns=cols)
    
    # append
    row = {
        "Data": action.get("Data", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        "Cod_Equip": action.get("Cod_Equip"),
        "DESCRICAO_EQUIPAMENTO": action.get("DESCRICAO_EQUIPAMENTO"),
        "Tipo_Servico": action.get("Tipo_Servico"),
        "Valor_Atual": action.get("Valor_Atual"),
        "Tipo_Controle": action.get("Tipo_Controle"),
        "Observacao": action.get("Observacao", ""),
        "Usuario": action.get("Usuario", "")
    }
    log_df = pd.concat([log_df, pd.DataFrame([row])], ignore_index=True)
    sheets["MANUT_LOG"] = log_df
    
    # grava tudo
    save_all_sheets(path, sheets)

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Controle Inteligente por Tipo")

    # Carrega dados
    df, df_frotas = load_data(EXCEL_PATH)

    # Inicializa st.session_state.thr de forma segura usando classes encontradas (evita KeyError)
    classes_found = []
    if "Classe_Operacional" in df.columns:
        classes_found = sorted(df["Classe_Operacional"].dropna().unique())
    elif "Classe_Operacional" in df_frotas.columns:
        classes_found = sorted(df_frotas["Classe_Operacional"].dropna().unique())

    if "thr" not in st.session_state:
        # padrão: min/max e intervalos km/hr
        st.session_state.thr = {}
        for cls in classes_found:
            st.session_state.thr[cls] = {"min": 1.5, "max": 5.0}

    # inicializa set para evitar gravações repetidas na sessão
    if "manut_processed" not in st.session_state:
        st.session_state.manut_processed = set()

    # Sidebar: tema e filtros e controles de manutenção
    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode (aplica visual escuro)", value=False)
        st.markdown("---")
        st.header("📅 Filtros")
        # Limpar filtros
        if st.button("🔄 Limpar Filtros"):
            st.session_state.clear()
            st.rerun()

        st.markdown("---")
        st.header("📈 Visual")
        top_n = st.slider("Número de categorias (Top N) antes de agrupar em 'Outros'", min_value=3, max_value=30, value=10)
        hide_text_threshold = st.slider("Esconder valores nas barras quando categorias >", min_value=5, max_value=40, value=8)

        st.markdown("---")
        st.header("🔧 Configuração de Intervalos")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Máquinas (HORAS)")
            INTERVALOS_TIPO['HORAS']['lubrificacao'] = st.number_input(
                "Lubrificação (h)", 
                min_value=10, max_value=2000, 
                value=INTERVALOS_TIPO['HORAS']['lubrificacao'], 
                step=10
            )
            INTERVALOS_TIPO['HORAS']['revisao'] = st.number_input(
                "Revisão (h)", 
                min_value=100, max_value=5000, 
                value=INTERVALOS_TIPO['HORAS']['revisao'], 
                step=50
            )
            
        with col2:
            st.subheader("Veículos (KM)")
            INTERVALOS_TIPO['QUILÔMETROS']['lubrificacao'] = st.number_input(
                "Lubrificação (km)", 
                min_value=1000, max_value=20000, 
                value=INTERVALOS_TIPO['QUILÔMETROS']['lubrificacao'], 
                step=500
            )
            INTERVALOS_TIPO['QUILÔMETROS']['revisao'] = st.number_input(
                "Revisão (km)", 
                min_value=5000, max_value=50000, 
                value=INTERVALOS_TIPO['QUILÔMETROS']['revisao'], 
                step=1000
            )

    # Aplica CSS leve
    apply_modern_css(dark_mode)

    # Paleta ativa
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"

    # Filtro (sem filtro de marca)
    def sidebar_filters_local(df: pd.DataFrame) -> dict:
        safra_opts = sorted(df["Safra"].dropna().unique()) if "Safra" in df.columns else []
        ano_opts = sorted(df["Ano"].dropna().unique()) if "Ano" in df.columns else []
        classe_opts = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []

        sel_safras = st.sidebar.multiselect("Safra", safra_opts, default=safra_opts[-1:] if safra_opts else [])
        sel_anos = st.sidebar.multiselect("Ano", ano_opts, default=ano_opts[-1:] if ano_opts else [])
        sel_meses = st.sidebar.multiselect("Mês (num)", sorted(df["Mes"].dropna().unique()) if "Mes" in df.columns else [], default=[datetime.now().month])
        st.sidebar.markdown("---")
        sel_classes = st.sidebar.multiselect("Classe Operacional", classe_opts, default=classe_opts)

        if not sel_safras:
            sel_safras = safra_opts
        if not sel_anos:
            sel_anos = ano_opts
        if not sel_meses:
            sel_meses = sorted(df["Mes"].dropna().unique()) if "Mes" in df.columns else []
        if not sel_classes:
            sel_classes = classe_opts

        return {
            "safras": sel_safras or [],
            "anos": sel_anos or [],
            "meses": sel_meses or [],
            "classes_op": sel_classes or [],
        }

    opts = sidebar_filters_local(df)
    df_f = filtrar_dados(df, opts)

    # Abas
    tab_principal, tab_consulta, tab_tabela, tab_config, tab_manut = st.tabs([
        "📊 Análise de Consumo",
        "🔎 Consulta de Frota",
        "📋 Tabela Detalhada",
        "⚙️ Configurações",
        "🛠️ Manutenção Inteligente"
    ])

    # ----- Aba Principal -----
    with tab_principal:
        if df_f.empty:
            st.warning("Sem dados para os filtros selecionados.")
            st.stop()

        kpis = calcular_kpis_consumo(df_f)
        total_eq = df_frotas.shape[0]
        ativos = int(df_frotas.query("ATIVO == 'ATIVO'").shape[0]) if "ATIVO" in df_frotas.columns else 0
        idade_media = (datetime.now().year - df_frotas["ANOMODELO"].median()) if "ANOMODELO" in df_frotas.columns else 0

        # NOVO: Mostrar distribuição por tipo de controle
        if "Tipo_Controle" in df_f.columns:
            tipo_dist = df_f["Tipo_Controle"].value_counts()
            maquinas_count = tipo_dist.get("HORAS", 0)
            veiculos_count = tipo_dist.get("QUILÔMETROS", 0)
        else:
            maquinas_count = 0
            veiculos_count = 0

        k1, k2, k3, k4, k5 = st.columns([1.4,1.4,1.2,1.2,1.2])
        with k1:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Litros Consumidos</div>'
                f'<div class="kpi-value">{formatar_brasileiro(kpis["total_litros"])}</div></div>',
                unsafe_allow_html=True
            )
        with k2:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Média de Consumo</div>'
                f'<div class="kpi-value">{formatar_brasileiro(kpis["media_consumo"])} km/l</div></div>',
                unsafe_allow_html=True
            )
        with k3:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Máquinas (Horas)</div>'
                f'<div class="kpi-value">{maquinas_count}</div></div>',
                unsafe_allow_html=True
            )
        with k4:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Veículos (KM)</div>'
                f'<div class="kpi-value">{veiculos_count}</div></div>',
                unsafe_allow_html=True
            )
        with k5:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Idade Média</div>'
                f'<div class="kpi-value">{idade_media:.0f} anos</div></div>',
                unsafe_allow_html=True
            )

        st.markdown("### ")
        st.info(f"🔍 {len(df_f):,} registros após aplicação dos filtros")

        # --- prepara df para plot com agrupamento Outras classes + top N logic ---
        df_plot_source = df_f.copy()
        df_plot_source["Classe_Operacional"] = df_plot_source["Classe_Operacional"].fillna("Sem Classe")
        # agrupa classes especificadas em OUTROS_CLASSES para limpeza inicial
        df_plot_source["Classe_Grouped"] = df_plot_source["Classe_Operacional"].apply(lambda s: "Outros" if s in OUTROS_CLASSES else s)

        # calcula média por classe agrupada
        media_op_full = df_plot_source.groupby("Classe_Grouped")["Media"].mean().reset_index()
        media_op_full["Media"] = media_op_full["Media"].round(1)

        # agora aplica top_n: manter top_n maiores por média, o resto vira "Outros"
        media_sorted = media_op_full.sort_values("Media", ascending=False).reset_index(drop=True)
        if media_sorted.shape[0] > top_n:
            top_keep = media_sorted.head(top_n)["Classe_Grouped"].tolist()
            # marca resto como Outros
            df_plot_source["Classe_TopN"] = df_plot_source["Classe_Grouped"].apply(lambda s: s if s in top_keep else "Outros")
            media_op = df_plot_source.groupby("Classe_TopN")["Media"].mean().reset_index().rename(columns={"Classe_TopN":"Classe_Grouped"})
            media_op["Media"] = media_op["Media"].round(1)
            outros_row = media_op[media_op["Classe_Grouped"] == "Outros"]
            media_op = media_op[media_op["Classe_Grouped"] != "Outros"].sort_values("Media", ascending=False)
            if not outros_row.empty:
                media_op = pd.concat([media_op, outros_row], ignore_index=True)
        else:
            media_op = media_sorted

        # wrapped labels
        media_op["Classe_wrapped"] = media_op["Classe_Grouped"].astype(str).apply(lambda s: wrap_labels(s, width=16))

        # plot
        hover_template_media = "Classe: %{x}<br>Média: %{y:.1f} km/l<extra></extra>"
        fig1 = make_bar(media_op, "Classe_wrapped", "Media",
                        "Média de Consumo por Classe Operacional",
                        {"Media": "Média (km/l)", "Classe_wrapped": "Classe"},
                        palette, rotate_x=-60, ticksize=10, height=520, hoverfmt=hover_template_media, wrap_width=16, hide_text_if_gt=hide_text_threshold)
        fig1.update_traces(marker_line_width=0.3)
        fig1.update_layout(template=plotly_template)
        st.plotly_chart(fig1, use_container_width=True, theme=None)

        # Gráfico 2 e 3 (consumo mensal e top10) mantidos — ajustando para esconder textos se necessário
        agg = df_f.groupby("AnoMes")["Qtde_Litros"].mean().reset_index()
        if not agg.empty:
            agg["Mes"] = pd.to_datetime(agg["AnoMes"] + "-01").dt.strftime("%b %Y")
            agg["Qtde_Litros"] = agg["Qtde_Litros"].round(1)
            hover_template_month = "Mês: %{x}<br>Litros: %{y:.1f} L<extra></extra>"
            fig2 = make_bar(agg, "Mes", "Qtde_Litros", "Consumo Mensal", {"Qtde_Litros": "Litros", "Mes": "Mês"}, palette, rotate_x=-45, ticksize=10, height=420, hoverfmt=hover_template_month, hide_text_if_gt=hide_text_threshold)
            fig2.update_layout(template=plotly_template)
            st.plotly_chart(fig2, use_container_width=True, theme=None)

        # Top10 equipamentos por Qtde_Litros total (mas mostra média de consumo)
        if "Cod_Equip" in df_f.columns and "Qtde_Litros" in df_f.columns:
            top10 = df_f.groupby("Cod_Equip")["Qtde_Litros"].sum().nlargest(10).index
            trend = (
                df_f[df_f["Cod_Equip"].isin(top10)]
                .groupby(["Cod_Equip", "Descricao_Equip"])["Media"].mean()
                .reset_index()
                .sort_values("Media", ascending=False)
            )
            if not trend.empty:
                trend["Equip_Label"] = trend.apply(lambda r: f"{r['Cod_Equip']} - {r['Descricao_Equip']}", axis=1)
                trend["Equip_Label_wrapped"] = trend["Equip_Label"].apply(lambda s: wrap_labels(s, width=18))
                trend["Media"] = trend["Media"].round(1)
                hover_template_top = "Equipamento: %{x}<br>Média: %{y:.1f} km/l<extra></extra>"
                fig3 = make_bar(trend, "Equip_Label_wrapped", "Media", "Média de Consumo por Equipamento (Top 10)",
                                {"Equip_Label_wrapped": "Equipamento", "Media": "Média (km/l)"},
                                palette, rotate_x=-45, ticksize=10, height=420, hoverfmt=hover_template_top, hide_text_if_gt=hide_text_threshold)
                fig3.update_traces(marker_line=dict(color="#000000", width=0.5))
                fig3.update_layout(template=plotly_template)
                st.plotly_chart(fig3, use_container_width=True, theme=None)

                # export fig3 if environment supports
                @st.cache_data(show_spinner=False)
                def get_fig_png(fig):
                    return fig.to_image(format="png", scale=2)

                try:
                    img = get_fig_png(fig3)
                    st.download_button("📷 Exportar Top10 (PNG)", data=img, file_name="top10.png", mime="image/png")
                except Exception:
                    st.caption("Exportação de imagem não disponível no ambiente atual.")

        # Gráfico 4 - consumo acumulado por safra
        st.markdown("---")
        st.header("📈 Comparativo de Consumo Acumulado por Safra")
        safras = sorted(df["Safra"].dropna().unique())
        sel_safras = st.multiselect("Selecione safras", safras, default=safras[-2:] if len(safras)>1 else safras)
        if sel_safras:
            df_cmp = df[df["Safra"].isin(sel_safras)].copy()
            iniciais = df_cmp.groupby("Safra")["Data"].min().to_dict()
            df_cmp["Dias_Uteis"] = (df_cmp["Data"] - df_cmp["Safra"].map(iniciais)).dt.days + 1
            df_cmp = df_cmp.groupby(["Safra", "Dias_Uteis"])["Qtde_Litros"].sum().groupby(level=0).cumsum().reset_index()
            hover_template_acum = "Dia: %{x}<br>Acumulado: %{y:.0f} L<extra></extra>"
            fig_acum = px.line(df_cmp, x="Dias_Uteis", y="Qtde_Litros", color="Safra", markers=True,
                               labels={"Dias_Uteis":"Dias desde início da safra","Qtde_Litros":"Consumo acumulado (L)"},
                               color_discrete_sequence=palette)
            fig_acum.update_layout(title="Consumo Acumulado por Safra", margin=dict(l=20,r=20,t=50,b=50), template=plotly_template, font=dict(size=13))
            fig_acum.update_traces(hovertemplate=hover_template_acum)
            ultima = sel_safras[-1]
            df_u = df_cmp[df_cmp["Safra"] == ultima]
            if not df_u.empty:
                fig_acum.add_scatter(x=[df_u["Dias_Uteis"].max()], y=[df_u["Qtde_Litros"].max()],
                                     mode="markers+text", text=[f"Hoje: {formatar_brasileiro(df_u['Qtde_Litros'].max())} L"],
                                     textposition="top right", marker=dict(size=8, color="#000000"), showlegend=False)
            st.plotly_chart(fig_acum, use_container_width=True, theme=None)

    # ----- Aba Consulta de Frota -----
    with tab_consulta:
        st.header("🔎 Ficha Individual do Equipamento")
        equip_label = st.selectbox("Selecione o Equipamento", options=df_frotas.sort_values("Cod_Equip")["label"])
        if equip_label:
            cod_sel = int(equip_label.split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod_sel").iloc[0]
            consumo_eq = df.query("Cod_Equip == @cod_sel").sort_values("Data", ascending=False)

            st.subheader(f"{dados_eq.get('DESCRICAO_EQUIPAMENTO','–')} ({dados_eq.get('PLACA','–')})")
            
            # NOVO: Mostra tipo de controle
            if not consumo_eq.empty:
                tipo_controle = consumo_eq.iloc[0].get("Tipo_Controle", "N/A")
                valor_atual = consumo_eq.iloc[0].get("Hod_Hor_Atual", np.nan)
                unidade = "h" if tipo_controle == "HORAS" else "km"
                valor_atual_display = f"{int(valor_atual)} {unidade}" if pd.notna(valor_atual) else "–"
            else:
                tipo_controle = "N/A"
                valor_atual_display = "–"
            
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Status", dados_eq.get("ATIVO", "–"))
            col2.metric("Placa", dados_eq.get("PLACA", "–"))
            col3.metric("Tipo Controle", tipo_controle)
            col4.metric("Valor Atual", valor_atual_display)
            col5.metric("Média Geral", formatar_brasileiro(consumo_eq["Media"].mean()))

            if not consumo_eq.empty:
                ultimo = consumo_eq.iloc[0]
                km_hs = ultimo.get("Km_Hs_Rod", np.nan)
                km_hs_display = str(int(km_hs)) if pd.notna(km_hs) else "–"
                safra_ult = consumo_eq["Safra"].max()
                df_safra = consumo_eq[consumo_eq["Safra"] == safra_ult]
                total_ult_safra = df_safra["Qtde_Litros"].sum()
                media_ult_safra = df_safra["Media"].mean()
            else:
                km_hs_display = "–"
                safra_ult = None
                total_ult_safra = None
                media_ult_safra = None

            c1, c2, c3, c4 = st.columns(4)
                c1.metric("Status", dados_eq.get("ATIVO", "–"))
                c2.metric("Placa", dados_eq.get("PLACA", "–"))
                c3.metric("Tipo de Controle", tipo_controle)
                c4.metric("Leitura Atual (Hod./Hor.)", valor_atual_display)


            st.markdown("---")
            st.subheader("Informações Cadastrais")
            st.dataframe(dados_eq.drop("label").to_frame("Valor"), use_container_width=True)

    # ----- Aba Tabela Detalhada -----
    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        cols = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO", 
                "Qtde_Litros", "Media", "Media_P", "Classe_Operacional", "Tipo_Controle", 
                "Hod_Hor_Atual", "Unid"]
        df_tab = df[[c for c in cols if c in df.columns]]

        csv_bytes = df_tab.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar CSV da Tabela", csv_bytes, "abastecimentos.csv", "text/csv")

        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        if "Media" in df_tab.columns:
            gb.configure_column("Media", type=["numericColumn"], precision=1)
        if "Qtde_Litros" in df_tab.columns:
            gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection("multiple", use_checkbox=True)
        gb.configure_side_bar()
        grid_options = gb.build()
        AgGrid(df_tab, gridOptions=grid_options, height=520, allow_unsafe_jscode=True)

    # ----- Aba Configurações -----
    with tab_config:
        st.header("⚙️ Configurações Avançadas")
        
        st.subheader("📊 Resumo da Frota por Tipo")
        if "Tipo_Controle" in df.columns:
            summary = df.groupby(["Tipo_Controle", "Classe_Operacional"]).size().reset_index(name="Quantidade")
            st.dataframe(summary, use_container_width=True)
        
        st.subheader("🔧 Padrões de Consumo por Classe")
        if "thr" not in st.session_state:
            classes = df["Classe_Operacional"].dropna().unique() if "Classe_Operacional" in df.columns else []
            st.session_state.thr = {cls: {"min": 1.5, "max": 5.0} for cls in classes}

        st.markdown("Personalize limites de consumo aceitável por classe:")
        for cls in sorted(st.session_state.thr.keys()):
            cols = st.columns(2)
            mn = cols[0].number_input(f"{cls} → Mínimo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["min"], step=0.1, key=f"min_{cls}")
            mx = cols[1].number_input(f"{cls} → Máximo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["max"], step=0.1, key=f"max_{cls}")
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx

        st.subheader("📋 Configuração de Intervalos por Tipo")
        st.markdown("Os intervalos abaixo são configurados na barra lateral e aplicados globalmente:")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"""
            **Máquinas (HORAS):**
            - Lubrificação: {INTERVALOS_TIPO['HORAS']['lubrificacao']}h
            - Revisão: {INTERVALOS_TIPO['HORAS']['revisao']}h
            """)
        with col2:
            st.info(f"""
            **Veículos (QUILÔMETROS):**
            - Lubrificação: {INTERVALOS_TIPO['QUILÔMETROS']['lubrificacao']}km
            - Revisão: {INTERVALOS_TIPO['QUILÔMETROS']['revisao']}km
            """)

    # ----- Aba Manutenção -----
    with tab_manut:
        st.header("🛠️ Controle Inteligente de Manutenção")
        st.markdown("Sistema diferenciado por tipo: **Máquinas** controladas por **horas**, **Veículos** por **quilômetros**.")
    
        # Gera tabela de manutenção
        mf = build_maintenance_table_new(df, df_frotas)
    
        # Estatísticas gerais
        # Filtra para não contar equipamentos sem dados de hodômetro/horímetro
        mf_valid = mf.dropna(subset=['Hod_Hor_Atual'])
        total_equipamentos_validos = len(mf_valid)
        com_alerta_lub = mf_valid["Alert_Lubrificacao"].sum()
        com_alerta_rev = mf_valid["Alert_Revisao"].sum()
        qualquer_alerta = mf_valid["Qualquer_Alerta"].sum()
    
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Equipamentos Monitorados", total_equipamentos_validos)
        col2.metric("Alertas Lubrificação", int(com_alerta_lub))
        col3.metric("Alertas Revisão", int(com_alerta_rev))
        col4.metric("Total com Alertas", int(qualquer_alerta))
    
        # Tabela: equipamentos com manutenção próxima ou vencida
        df_due = mf[mf["Qualquer_Alerta"]].copy().sort_values(["Para_Revisao", "Para_Lubrificacao"], ascending=[True, True])
    
        st.subheader("⚠️ Equipamentos com Manutenção Próxima/Atrasada")
        
        if not df_due.empty:
            # Organiza por tipo de alerta para exibição
            lub_only = df_due[(df_due["Alert_Lubrificacao"]) & (~df_due["Alert_Revisao"])]
            rev_only = df_due[(~df_due["Alert_Lubrificacao"]) & (df_due["Alert_Revisao"])]
            both_alerts = df_due[(df_due["Alert_Lubrificacao"]) & (df_due["Alert_Revisao"])]
    
            if not both_alerts.empty:
                st.markdown("##### 🔴 **Crítico: Lubrificação E Revisão**")
                display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Hod_Hor_Atual", "Para_Lubrificacao", "Para_Revisao", "Unidade"]
                st.dataframe(both_alerts[[c for c in display_cols if c in both_alerts.columns]].reset_index(drop=True), use_container_width=True)
    
            if not rev_only.empty:
                st.markdown("##### 🟡 **Revisão Próxima**")
                display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Hod_Hor_Atual", "Para_Revisao", "Unidade"]
                st.dataframe(rev_only[[c for c in display_cols if c in rev_only.columns]].reset_index(drop=True), use_container_width=True)
    
            if not lub_only.empty:
                st.markdown("##### 🔵 **Lubrificação Próxima**")
                display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Hod_Hor_Atual", "Para_Lubrificacao", "Unidade"]
                st.dataframe(lub_only[[c for c in display_cols if c in lub_only.columns]].reset_index(drop=True), use_container_width=True)
    
            # export CSV
            all_display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Controle", "Hod_Hor_Atual", "Para_Lubrificacao", "Para_Revisao", "Alert_Lubrificacao", "Alert_Revisao", "Unidade"]
            csvm = df_due[[c for c in all_display_cols if c in df_due.columns]].to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Exportar CSV - Equipamentos em alerta", csvm, "manutencao_alerta.csv", "text/csv")
        else:
            st.success("✅ Nenhum equipamento com alerta de manutenção dentro dos parâmetros configurados!")
    
        st.markdown("---")
        st.subheader("✅ Registrar Manutenção Realizada")
        
        if not df_due.empty:
            for _, row in df_due.iterrows():
                cod = int(row["Cod_Equip"]) if not pd.isna(row["Cod_Equip"]) else None
                label = f"{int(cod)} - {row.get('DESCRICAO_EQUIPAMENTO','')}" if cod else str(row.get('DESCRICAO_EQUIPAMENTO',''))
                
                with st.expander(f"🔧 {label}"):
                    cols = st.columns([2,1,1])
                    
                    tipo_controle = row.get("Tipo_Controle", "N/A")
                    valor_atual = row.get("Hod_Hor_Atual", np.nan)
                    unidade = row.get("Unidade", "")
                    
                    valor_display = valor_atual if pd.notna(valor_atual) else 0
                    
                    cols[0].markdown(f"**Tipo:** {tipo_controle}")
                    cols[1].markdown(f"**Atual:** {valor_display:.0f} {unidade}")
                    
                    servico_options = []
                    if row.get("Alert_Lubrificacao", False): servico_options.append("Lubrificação")
                    if row.get("Alert_Revisao", False): servico_options.append("Revisão")
                    
                    if servico_options:
                        servico = st.selectbox(f"Tipo de serviço realizado", servico_options, key=f"servico_{cod}")
                        observacao = st.text_input(f"Observações", key=f"obs_{cod}")
                        
                        if st.button(f"✅ Registrar {servico}", key=f"reg_{cod}"):
                            key = f"manut_done_{cod}_{servico}"
                            if key not in st.session_state.get('manut_processed', set()):
                                action = {
                                    "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Cod_Equip": cod,
                                    "DESCRICAO_EQUIPAMENTO": row.get("DESCRICAO_EQUIPAMENTO", ""),
                                    "Tipo_Servico": servico,
                                    "Valor_Atual": float(valor_atual) if pd.notna(valor_atual) else np.nan,
                                    "Tipo_Controle": tipo_controle,
                                    "Observacao": observacao,
                                    "Usuario": "usuario_app"
                                }
                                try:
                                    append_manut_log_new(EXCEL_PATH, action)
                                    st.success(f"✅ {servico} registrada para equipamento {cod}!")
                                    st.session_state.setdefault('manut_processed', set()).add(key)
                                    # --- LINHA ADICIONADA PARA ATUALIZAÇÃO AUTOMÁTICA ---
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Falha ao registrar manutenção: {e}")
                            else:
                                st.warning("Já registrado nesta sessão.")
        else:
            st.info("Nenhum equipamento em alerta para registrar manutenção.")
    
        st.markdown("---")
        st.subheader("📊 Visão Geral da Frota (Plano de Manutenção)")
        
        mf_sorted = mf.sort_values(["Qualquer_Alerta", "Para_Revisao"], ascending=[False, True])
        
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Controle", "Hod_Hor_Atual", "Prox_Lubrificacao", "Para_Lubrificacao", "Prox_Revisao", "Para_Revisao", "Unidade"]
        mf_display = mf_sorted[[c for c in overview_cols if c in mf_sorted.columns]].dropna(subset=['Hod_Hor_Atual'])
        
        st.dataframe(mf_display.reset_index(drop=True), use_container_width=True)
        
        csv_over = mf_display.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar CSV - Plano de Manutenção Completo", csv_over, "manutencao_completa.csv", "text/csv")

if __name__ == "__main__":
    main()
