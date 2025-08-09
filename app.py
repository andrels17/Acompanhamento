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
    def inferir_tipo_por_classe(row):
        if pd.notna(row['Tipo_Controle']):
            return row['Tipo_Controle']
        
        classe = str(row.get('Classe_Operacional', '')).upper()
        
        # Máquinas agrícolas geralmente são por HORAS
        if any(palavra in classe for palavra in ['TRATOR', 'COLHEITADEIRA', 'PULVERIZADOR', 'PLANTADEIRA', 'PÁ CARREGADEIRA']):
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

    # Normaliza abastecimento (mantendo nomes originais)
    df_abast.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
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

    # Marca / Fazenda
    df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    df["Fazenda"] = df["Ref1"].astype(str)

    # Cálculo seguro de Consumo km/l
    if "Km_Hs_Rod" in df.columns and "Qtde_Litros" in df.columns:
        df["Consumo_km_l"] = np.where(df["Qtde_Litros"] > 0, df["Km_Hs_Rod"] / df["Qtde_Litros"], np.nan)
        df["Media"] = df["Consumo_km_l"]
    else:
        df["Consumo_km_l"] = np.nan

    return df, df_frotas

@st.cache_data
def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Filtra o DataFrame conforme opções selecionadas."""
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
    current_data = get_current_value_from_bd(df_abast)
    mf = df_frotas.copy()
    mf = mf.merge(current_data, left_on="Cod_Equip", right_index=True, how="left")
    
    def calcular_proximos_servicos(row):
        tipo_controle = row.get("Tipo_Controle")
        valor_atual = row.get("Hod_Hor_Atual")
        
        if pd.isna(valor_atual) or pd.isna(tipo_controle) or tipo_controle not in INTERVALOS_TIPO:
            return pd.Series({
                "Prox_Lubrificacao": np.nan, "Prox_Revisao": np.nan,
                "Para_Lubrificacao": np.nan, "Para_Revisao": np.nan,
                "Alert_Lubrificacao": False, "Alert_Revisao": False, "Unidade": ""
            })

        intervalos = INTERVALOS_TIPO[tipo_controle]
        alertas = ALERTAS_TIPO[tipo_controle]
        
        prox_lub = ((valor_atual // intervalos["lubrificacao"]) + 1) * intervalos["lubrificacao"]
        prox_rev = ((valor_atual // intervalos["revisao"]) + 1) * intervalos["revisao"]
        
        para_lub = prox_lub - valor_atual
        para_rev = prox_rev - valor_atual
        
        alert_lub = para_lub <= alertas["lubrificacao"]
        alert_rev = para_rev <= alertas["revisao"]
        
        unidade = "h" if tipo_controle == "HORAS" else "km"
        
        return pd.Series({
            "Prox_Lubrificacao": prox_lub, "Prox_Revisao": prox_rev,
            "Para_Lubrificacao": para_lub, "Para_Revisao": para_rev,
            "Alert_Lubrificacao": alert_lub, "Alert_Revisao": alert_rev, "Unidade": unidade
        })

    maintenance_data = mf.apply(calcular_proximos_servicos, axis=1)
    mf = pd.concat([mf, maintenance_data], axis=1)
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
            for name, df_sheet in sheets.items():
                df_sheet.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar o Excel: {e}")
        raise

def append_manut_log_new(path: str, action: dict):
    """Adiciona uma linha de log de manutenção e salva no Excel."""
    sheets = read_all_sheets(path)
    if sheets is None:
        sheets = {}
    
    log_cols = ["Data", "Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Servico", "Valor_Atual", "Tipo_Controle", "Observacao", "Usuario"]
    log_df = sheets.get("MANUT_LOG", pd.DataFrame(columns=log_cols))
    
    log_df = pd.concat([log_df, pd.DataFrame([action])], ignore_index=True)
    sheets["MANUT_LOG"] = log_df
    
    save_all_sheets(path, sheets)

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Controle Inteligente por Tipo")

    df, df_frotas = load_data(EXCEL_PATH)

    if "manut_processed" not in st.session_state:
        st.session_state.manut_processed = set()

    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode", value=True)
        st.markdown("---")
        st.header("📅 Filtros")
        if st.button("🔄 Limpar Filtros"):
            st.session_state.clear()
            st.rerun()

        # Filtros
        safra_opts = sorted(df["Safra"].dropna().unique()) if "Safra" in df.columns else []
        ano_opts = sorted(df["Ano"].dropna().unique()) if "Ano" in df.columns else []
        mes_opts = sorted(df["Mes"].dropna().unique()) if "Mes" in df.columns else []
        classe_opts = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []

        sel_safras = st.multiselect("Safra", safra_opts, default=safra_opts[-1:] if safra_opts else [])
        sel_anos = st.multiselect("Ano", ano_opts, default=ano_opts[-1:] if ano_opts else [])
        sel_meses = st.multiselect("Mês", mes_opts, default=[datetime.now().month])
        sel_classes = st.multiselect("Classe Operacional", classe_opts, default=classe_opts)

        opts = {
            "safras": sel_safras or safra_opts,
            "anos": sel_anos or ano_opts,
            "meses": sel_meses or mes_opts,
            "classes_op": sel_classes or classe_opts,
        }

        st.markdown("---")
        st.header("🔧 Configuração de Intervalos")
        with st.expander("Intervalos de Manutenção", expanded=False):
            st.subheader("Máquinas (HORAS)")
            INTERVALOS_TIPO['HORAS']['lubrificacao'] = st.number_input("Lubrificação (h)", min_value=10, max_value=2000, value=INTERVALOS_TIPO['HORAS']['lubrificacao'], step=10, key="h_lub")
            INTERVALOS_TIPO['HORAS']['revisao'] = st.number_input("Revisão (h)", min_value=100, max_value=5000, value=INTERVALOS_TIPO['HORAS']['revisao'], step=50, key="h_rev")
            
            st.subheader("Veículos (KM)")
            INTERVALOS_TIPO['QUILÔMETROS']['lubrificacao'] = st.number_input("Lubrificação (km)", min_value=1000, max_value=20000, value=INTERVALOS_TIPO['QUILÔMETROS']['lubrificacao'], step=500, key="km_lub")
            INTERVALOS_TIPO['QUILÔMETROS']['revisao'] = st.number_input("Revisão (km)", min_value=5000, max_value=50000, value=INTERVALOS_TIPO['QUILÔMETROS']['revisao'], step=1000, key="km_rev")

    apply_modern_css(dark_mode)
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"
    
    df_f = filtrar_dados(df, opts)

    tab_principal, tab_consulta, tab_tabela, tab_manut = st.tabs([
        "📊 Análise de Consumo",
        "🔎 Consulta de Frota",
        "📋 Tabela Detalhada",
        "🛠️ Manutenção Inteligente"
    ])

    with tab_principal:
        if df_f.empty:
            st.warning("Sem dados para os filtros selecionados.")
        else:
            kpis = calcular_kpis_consumo(df_f)
            idade_media = (datetime.now().year - df_frotas["ANOMODELO"].median()) if "ANOMODELO" in df_frotas.columns else 0

            tipo_dist = df_f["Tipo_Controle"].value_counts()
            maquinas_count = tipo_dist.get("HORAS", 0)
            veiculos_count = tipo_dist.get("QUILÔMETROS", 0)

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.markdown(f'<div class="kpi-card"><div class="kpi-title">Litros Consumidos</div><div class="kpi-value">{formatar_brasileiro(kpis["total_litros"])}</div></div>', unsafe_allow_html=True)
            k2.markdown(f'<div class="kpi-card"><div class="kpi-title">Média Consumo</div><div class="kpi-value">{formatar_brasileiro(kpis["media_consumo"])} km/l</div></div>', unsafe_allow_html=True)
            k3.markdown(f'<div class="kpi-card"><div class="kpi-title">Máquinas (Horas)</div><div class="kpi-value">{maquinas_count}</div></div>', unsafe_allow_html=True)
            k4.markdown(f'<div class="kpi-card"><div class="kpi-title">Veículos (KM)</div><div class="kpi-value">{veiculos_count}</div></div>', unsafe_allow_html=True)
            k5.markdown(f'<div class="kpi-card"><div class="kpi-title">Idade Média</div><div class="kpi-value">{idade_media:.0f} anos</div></div>', unsafe_allow_html=True)

    # ----- Aba Consulta de Frota (VERSÃO CORRIGIDA E COMPLETA) -----
    with tab_consulta:
        st.header("🔎 Ficha Individual do Equipamento")
        equip_options = df_frotas.sort_values("Cod_Equip")["label"]
        equip_label = st.selectbox("Selecione o Equipamento", options=equip_options, key="select_equip_consulta")
        
        if equip_label:
            cod_sel = int(equip_label.split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod_sel").iloc[0]
            consumo_eq = df.query("Cod_Equip == @cod_sel").sort_values("Data", ascending=False)
    
            st.subheader(f"{dados_eq.get('DESCRICAO_EQUIPAMENTO','–')} ({dados_eq.get('PLACA','–')})")
            
            if not consumo_eq.empty:
                ultimo_registro = consumo_eq.iloc[0]
                tipo_controle = ultimo_registro.get("Tipo_Controle", "N/A")
                valor_atual_leitura = ultimo_registro.get("Hod_Hor_Atual", np.nan)
                unidade = "h" if tipo_controle == "HORAS" else "km"
                valor_atual_display = f"{int(valor_atual_leitura):,} {unidade}".replace(",",".") if pd.notna(valor_atual_leitura) else "–"
                
                media_geral_eq = consumo_eq["Media"].mean()
                total_consumido_eq = consumo_eq["Qtde_Litros"].sum()
                safra_ult = consumo_eq["Safra"].max()
                df_safra = consumo_eq[consumo_eq["Safra"] == safra_ult]
                total_ult_safra = df_safra["Qtde_Litros"].sum()
                media_ult_safra = df_safra["Media"].mean()
            else:
                tipo_controle, valor_atual_display, media_geral_eq, total_consumido_eq, safra_ult, total_ult_safra, media_ult_safra = "N/A", "–", 0, 0, "", 0, 0
    
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Status", dados_eq.get("ATIVO", "–"))
            col2.metric("Placa", dados_eq.get("PLACA", "–"))
            col3.metric("Média Geral", formatar_brasileiro(media_geral_eq))
            col4.metric("Total Consumido (L)", formatar_brasileiro(total_consumido_eq))
    
            st.markdown("---") 
    
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("Tipo de Controle", tipo_controle)
            c6.metric("Leitura Atual (Hod./Hor.)", valor_atual_display)
            c7.metric(f"Total Última Safra {f'({safra_ult})' if safra_ult else ''}", formatar_brasileiro(total_ult_safra))
            c8.metric("Média Última Safra", formatar_brasileiro(media_ult_safra))
    
            st.markdown("---")
            st.subheader("Informações Cadastrais")
            st.dataframe(dados_eq.drop("label", errors='ignore').to_frame("Valor"), use_container_width=True)

    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        cols_to_show = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "ANOMODELO", "Qtde_Litros", "Media", "Classe_Operacional", "Tipo_Controle", "Hod_Hor_Atual", "Unid"]
        df_tab = df_f[[c for c in cols_to_show if c in df_f.columns]]

        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
        gb.configure_side_bar()
        AgGrid(df_tab, gridOptions=gb.build(), height=600, allow_unsafe_jscode=True, theme='streamlit' if not dark_mode else 'alpine')

    with tab_manut:
        st.header("🛠️ Controle Inteligente de Manutenção")
        
        mf = build_maintenance_table_new(df, df_frotas)

        com_alerta_lub = mf["Alert_Lubrificacao"].sum()
        com_alerta_rev = mf["Alert_Revisao"].sum()
        qualquer_alerta = mf["Qualquer_Alerta"].sum()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Equipamentos", len(mf.dropna(subset=['Hod_Hor_Atual'])))
        col2.metric("Alertas Lubrificação", int(com_alerta_lub))
        col3.metric("Alertas Revisão", int(com_alerta_rev))
        col4.metric("Total com Alertas", int(qualquer_alerta))

        st.markdown("---")
        st.subheader("⚠️ Equipamentos com Manutenção Próxima/Atrasada")
        
        df_due = mf.loc[mf["Qualquer_Alerta"] == True].copy()
        
        if not df_due.empty:
            display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Controle", "Hod_Hor_Atual", "Para_Lubrificacao", "Para_Revisao", "Unidade"]
            st.dataframe(df_due[[c for c in display_cols if c in df_due.columns]].reset_index(drop=True))
        else:
            st.success("✅ Nenhum equipamento com alerta de manutenção dentro dos parâmetros configurados!")

        st.markdown("---")
        st.subheader("✅ Registrar Manutenção Realizada")
        
        if not df_due.empty:
            for _, row in df_due.iterrows():
                cod = int(row["Cod_Equip"])
                label = f"{cod} - {row.get('DESCRICAO_EQUIPAMENTO','')}"
                with st.expander(f"🔧 {label}"):
                    servico_options = []
                    if row.get("Alert_Lubrificacao"): servico_options.append("Lubrificação")
                    if row.get("Alert_Revisao"): servico_options.append("Revisão")
                    
                    if servico_options:
                        servico = st.selectbox("Tipo de serviço", servico_options, key=f"serv_{cod}")
                        obs = st.text_input("Observações", key=f"obs_{cod}")
                        if st.button("Registrar", key=f"reg_{cod}"):
                            action_key = f"manut_done_{cod}_{servico}"
                            if action_key not in st.session_state.manut_processed:
                                action = {
                                    "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Cod_Equip": cod,
                                    "DESCRICAO_EQUIPAMENTO": row.get("DESCRICAO_EQUIPAMENTO"), "Tipo_Servico": servico,
                                    "Valor_Atual": row.get("Hod_Hor_Atual"), "Tipo_Controle": row.get("Tipo_Controle"),
                                    "Observacao": obs, "Usuario": "usuario_app"
                                }
                                try:
                                    append_manut_log_new(EXCEL_PATH, action)
                                    st.success(f"Manutenção de {servico} registrada para {label}!")
                                    st.session_state.manut_processed.add(action_key)
                                except Exception as e:
                                    st.error(f"Falha ao registrar: {e}")
                            else:
                                st.warning("Já registrado nesta sessão.")
        else:
            st.info("Nenhum equipamento em alerta para registrar manutenção.")

        st.markdown("---")
        st.subheader("📊 Visão Geral da Frota")
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo_Controle", "Hod_Hor_Atual", "Prox_Lubrificacao", "Para_Lubrificacao", "Prox_Revisao", "Para_Revisao", "Unidade"]
        mf_display = mf[[c for c in overview_cols if c in mf.columns]].dropna(subset=['Hod_Hor_Atual'])
        st.dataframe(mf_display.sort_values(["Qualquer_Alerta", "Para_Revisao"], ascending=[False, True]), use_container_width=True)

if __name__ == "__main__":
    main()
