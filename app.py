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

# Possíveis nomes de colunas para hodômetro e horímetro — atualize se necessário
HODOMETRO_COLS = ["HODOMETRO", "Hodometro", "Km_Atual", "KM", "Km", "KmAtual", "Km_Atual", "Km_Ultima_Manutencao"]
HORIMETRO_COLS = ["HORIMETRO", "Horimetro", "Hr_Atual", "Horas", "Horimetro_Horas", "Horimetros", "Hr_Ultima_Manutencao"]

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

def find_first_column(df: pd.DataFrame, candidates: list) -> str | None:
    """Retorna o primeiro nome de coluna existente em df a partir da lista de candidatos."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

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
        "Potencia_CV_Abast"
    ]

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
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod"]:
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
        </style>
        """,
        unsafe_allow_html=True
    )

# ---------------- Manutenção: lógica ----------------
def detect_odometer_and_hourmeter(df_frotas: pd.DataFrame, df_abast: pd.DataFrame):
    """Encontra colunas reais usadas para hodômetro/horímetro e constrói colunas consolidadas."""
    hod_col = find_first_column(df_frotas, HODOMETRO_COLS)
    hr_col = find_first_column(df_frotas, HORIMETRO_COLS)

    # Se não achar em frotas, tenta extrair do histórico (abastecimento) usando Km_Hs_Rod as fallback
    if hod_col is None and "Km_Hs_Rod" in df_abast.columns:
        # pega o último registro por equipamento como estimativa
        last_km = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Km_Hs_Rod"].last().rename("Km_current_from_abast")
        return None, hr_col, last_km
    return hod_col, hr_col, None

def build_maintenance_table(df_frotas: pd.DataFrame, last_km_series: pd.Series | None,
                            km_interval_default: int, hr_interval_default: int,
                            class_intervals: dict):
    """Constrói tabela com próximos serviços/lubrificação."""
    mf = df_frotas.copy()
    hod_col = find_first_column(mf, HODOMETRO_COLS)
    hr_col = find_first_column(mf, HORIMETRO_COLS)

    if hod_col and hod_col in mf.columns:
        mf["Km_Current"] = pd.to_numeric(mf[hod_col], errors="coerce")
    elif last_km_series is not None:
        mf = mf.set_index("Cod_Equip")
        mf["Km_Current"] = last_km_series.reindex(mf.index)
        mf = mf.reset_index()
    else:
        # tenta coluna Hodometro_Atual em BD? (não aqui), default NaN
        mf["Km_Current"] = np.nan

    if hr_col and hr_col in mf.columns:
        mf["Hr_Current"] = pd.to_numeric(mf[hr_col], errors="coerce")
    else:
        mf["Hr_Current"] = np.nan

    # Se existir colunas de última manutenção, usa-as
    if "Km_Ultima_Manutencao" in mf.columns:
        mf["Km_Last_Service"] = pd.to_numeric(mf["Km_Ultima_Manutencao"], errors="coerce")
    else:
        mf["Km_Last_Service"] = np.nan

    if "Hr_Ultima_Manutencao" in mf.columns:
        mf["Hr_Last_Service"] = pd.to_numeric(mf["Hr_Ultima_Manutencao"], errors="coerce")
    else:
        mf["Hr_Last_Service"] = np.nan

    # define intervalos por equipamento (classe)
    def get_interval(row, kind):
        cls = row.get("Classe_Operacional", "")
        if cls in class_intervals and class_intervals[cls].get(kind) is not None:
            return class_intervals[cls].get(kind)
        return km_interval_default if kind == "km" else hr_interval_default

    mf["Km_Service_Interval"] = mf.apply(lambda r: get_interval(r, "km"), axis=1)
    mf["Hr_Service_Interval"] = mf.apply(lambda r: get_interval(r, "hr"), axis=1)

    # cálculo do próximo serviço, preferindo last service quando disponível
    mf["Km_Next_Service"] = np.where(
        mf["Km_Last_Service"].notna(),
        mf["Km_Last_Service"] + mf["Km_Service_Interval"],
        np.where(mf["Km_Current"].notna(), mf["Km_Current"] + mf["Km_Service_Interval"], np.nan)
    )
    mf["Km_To_Service"] = mf["Km_Next_Service"] - mf["Km_Current"]

    mf["Hr_Next_Oil"] = np.where(
        mf["Hr_Last_Service"].notna(),
        mf["Hr_Last_Service"] + mf["Hr_Service_Interval"],
        np.where(mf["Hr_Current"].notna(), mf["Hr_Current"] + mf["Hr_Service_Interval"], np.nan)
    )
    mf["Hr_To_Oil"] = mf["Hr_Next_Oil"] - mf["Hr_Current"]

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

def append_manut_log(path: str, action: dict):
    """
    action: dict com keys:
    - Cod_Equip, DESCRICAO_EQUIPAMENTO, Tipo (KM/HR/BOTH), Km_Current, Hr_Current, Intervalo_KM, Intervalo_HR, Observacao, Data
    """
    sheets = read_all_sheets(path)
    if sheets is None:
        sheets = {}
    log_df = None
    if "MANUT_LOG" in sheets:
        log_df = sheets["MANUT_LOG"]
    else:
        # cria colunas básicas
        cols = ["Data", "Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Tipo", "Km_Current", "Hr_Current", "Intervalo_KM", "Intervalo_HR", "Observacao", "Usuario"]
        log_df = pd.DataFrame(columns=cols)
    # append
    row = {
        "Data": action.get("Data", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        "Cod_Equip": action.get("Cod_Equip"),
        "DESCRICAO_EQUIPAMENTO": action.get("DESCRICAO_EQUIPAMENTO"),
        "Tipo": action.get("Tipo"),
        "Km_Current": action.get("Km_Current"),
        "Hr_Current": action.get("Hr_Current"),
        "Intervalo_KM": action.get("Intervalo_KM"),
        "Intervalo_HR": action.get("Intervalo_HR"),
        "Observacao": action.get("Observacao", ""),
        "Usuario": action.get("Usuario", "")
    }
    log_df = pd.concat([log_df, pd.DataFrame([row])], ignore_index=True)
    sheets["MANUT_LOG"] = log_df

    # Também atualiza a aba FROTAS (se existir) com Km_Ultima_Manutencao / Hr_Ultima_Manutencao
    # Se não existir FROTAS, tentamos BD.
    target_sheet = "FROTAS" if "FROTAS" in sheets else "BD" if "BD" in sheets else None
    if target_sheet:
        df_target = sheets[target_sheet].copy()
        # ajustar nomes de colunas: se já houver 'Cod_Equip' manter, caso diferente tentar mapear
        # Supondo que a aba FROTAS já tem coluna Cod_Equip
        if "Cod_Equip" not in df_target.columns and "COD_EQUIPAMENTO" in df_target.columns:
            df_target = df_target.rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"})
        # localiza linha por Cod_Equip
        cod = action.get("Cod_Equip")
        mask = df_target["Cod_Equip"] == cod
        if mask.any():
            if "Km_Ultima_Manutencao" in df_target.columns:
                df_target.loc[mask, "Km_Ultima_Manutencao"] = action.get("Km_Current")
            else:
                # cria coluna
                df_target["Km_Ultima_Manutencao"] = pd.NA
                df_target.loc[mask, "Km_Ultima_Manutencao"] = action.get("Km_Current")
            if "Hr_Ultima_Manutencao" in df_target.columns:
                df_target.loc[mask, "Hr_Ultima_Manutencao"] = action.get("Hr_Current")
            else:
                df_target["Hr_Ultima_Manutencao"] = pd.NA
                df_target.loc[mask, "Hr_Ultima_Manutencao"] = action.get("Hr_Current")
            sheets[target_sheet] = df_target
        else:
            # se não achou por Cod_Equip, apenas atualiza a planilha com log (não quebra)
            sheets[target_sheet] = df_target

    # grava tudo
    save_all_sheets(path, sheets)

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Visual Moderno (Light Premium)")

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
            st.session_state.thr[cls] = {"min": 1.5, "max": 5.0, "km_interval": 10000, "hr_interval": 250}

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
        st.header("🔧 Manutenção & Lubrificação")
        st.markdown("Defina intervalos padrão (por km e por horas). Você pode também definir intervalos por classe na aba Configurações.")
        km_interval_default = st.number_input("Intervalo padrão (km) para revisão", min_value=100, max_value=200000, value=10000, step=100)
        hr_interval_default = st.number_input("Intervalo padrão (horas) para lubrificação", min_value=1, max_value=5000, value=250, step=1)
        km_due_threshold = st.number_input("Alerta para revisão se faltar <= (km)", min_value=10, max_value=5000, value=500, step=10)
        hr_due_threshold = st.number_input("Alerta para lubrificação se faltar <= (horas)", min_value=1, max_value=500, value=20, step=1)

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
        "🛠️ Manutenção"
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

        k1, k2, k3, k4 = st.columns([1.6,1.6,1.4,1.4])
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
                '<div class="kpi-card"><div class="kpi-title">Veículos Ativos</div>'
                f'<div class="kpi-value">{ativos} / {total_eq}</div></div>',
                unsafe_allow_html=True
            )
        with k4:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Idade Média da Frota</div>'
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
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Status", dados_eq.get("ATIVO", "–"))
            col2.metric("Placa", dados_eq.get("PLACA", "–"))
            col3.metric("Média Geral", formatar_brasileiro(consumo_eq["Media"].mean()))
            col4.metric("Total Consumido (L)", formatar_brasileiro(consumo_eq["Qtde_Litros"].sum()))

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

            c5, c6, c7 = st.columns(3)
            c5.metric("KM/Hr Último Registro", km_hs_display)
            c6.metric(f"Total Última Safra{f' ({safra_ult})' if safra_ult else ''}", formatar_brasileiro(total_ult_safra) if total_ult_safra is not None else "–")
            c7.metric("Média Última Safra", formatar_brasileiro(media_ult_safra) if media_ult_safra is not None else "–")

            st.markdown("---")
            st.subheader("Informações Cadastrais")
            st.dataframe(dados_eq.drop("label").to_frame("Valor"), use_container_width=True)

    # ----- Aba Tabela Detalhada -----
    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        cols = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO", "Qtde_Litros", "Media", "Media_P", "Classe_Operacional"]
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
        st.header("⚙️ Padrões por Classe Operacional (Alertas)")
        if "thr" not in st.session_state:
            classes = df["Classe_Operacional"].dropna().unique() if "Classe_Operacional" in df.columns else []
            st.session_state.thr = {cls: {"min": 1.5, "max": 5.0, "km_interval": 10000, "hr_interval": 250} for cls in classes}

        st.markdown("Personalize intervalo de manutenção por classe (opcional):")
        for cls in sorted(st.session_state.thr.keys()):
            cols = st.columns(3)
            mn = cols[0].number_input(f"{cls} → Mínimo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["min"], step=0.1, key=f"min_{cls}")
            mx = cols[1].number_input(f"{cls} → Máximo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["max"], step=0.1, key=f"max_{cls}")
            kint = cols[2].number_input(f"{cls} → Intervalo revisão (km)", min_value=0, max_value=200000, value=st.session_state.thr[cls]["km_interval"], step=100, key=f"kmint_{cls}")
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx
            st.session_state.thr[cls]["km_interval"] = int(kint)

    # ----- Aba Manutenção -----
    with tab_manut:
        st.header("🛠️ Controle de Revisões e Lubrificação")
        st.markdown("O sistema tenta identificar hodômetros/horímetros e calcular próximos serviços com base em intervalos padrão ou por classe.")

        # detect colunas reais (e uma série fallback do histórico)
        hod_col, hr_col, last_km_series = detect_odometer_and_hourmeter(df_frotas, df)

        if hod_col:
            st.markdown(f"**Hodômetro encontrado em frotas:** `{hod_col}`")
        elif last_km_series is not None:
            st.markdown("**Hodômetro:** não encontrado diretamente nas colunas de frotas; usando histórico como fallback (Km_Hs_Rod).")
        else:
            st.markdown("**Hodômetro:** não encontrado.")

        if hr_col:
            st.markdown(f"**Horímetro encontrado em frotas:** `{hr_col}`")
        else:
            st.markdown("**Horímetro:** não encontrado.")

        # montar dict de intervalos por classe a partir de st.session_state.thr, se existir
        class_intervals = {}
        if "thr" in st.session_state:
            for cls, v in st.session_state.thr.items():
                class_intervals[cls] = {"km": v.get("km_interval", None), "hr": v.get("hr_interval", None)}

        mf = build_maintenance_table(df_frotas, last_km_series, int(km_interval_default), int(hr_interval_default), class_intervals)

        # calcula flags de proximidade de manutenção
        mf["Km_To_Service"] = mf["Km_Next_Service"] - mf["Km_Current"]
        mf["Hr_To_Oil"] = mf["Hr_Next_Oil"] - mf["Hr_Current"]

        mf["Due_Km"] = mf["Km_To_Service"].apply(lambda x: True if pd.notna(x) and x <= km_due_threshold else False)
        mf["Due_Hr"] = mf["Hr_To_Oil"].apply(lambda x: True if pd.notna(x) and x <= hr_due_threshold else False)

        mf["Any_Due"] = mf["Due_Km"] | mf["Due_Hr"]

        # Tabela: equipamentos com manutenção próxima ou vencida
        df_due = mf[mf["Any_Due"]].copy().sort_values(["Due_Km", "Due_Hr"], ascending=False)

        st.subheader("Equipamentos com manutenção próxima/atrasada")
        st.write(f"Total equipamentos com alerta: {len(df_due)}")
        if not df_due.empty:
            display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Km_Current", "Km_Next_Service", "Km_To_Service", "Hr_Current", "Hr_Next_Oil", "Hr_To_Oil", "Due_Km", "Due_Hr"]
            available = [c for c in display_cols if c in df_due.columns]
            st.dataframe(df_due[available].reset_index(drop=True), use_container_width=True)

            # export CSV
            csvm = df_due[available].to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Exportar CSV - Equipamentos em alerta", csvm, "manutencao_alerta.csv", "text/csv")
        else:
            st.info("Nenhum equipamento com alerta de manutenção dentro dos thresholds configurados.")

        st.markdown("---")
        st.subheader("Marcar manutenção realizada")
        st.markdown("Marque a manutenção/lubrificação realizada — isso criará um registro em `MANUT_LOG` na planilha e atualizará a coluna de última manutenção na aba FROTAS/BD.")

        if not df_due.empty:
            # lista linhas com checkboxes
            for _, row in df_due.iterrows():
                cod = int(row["Cod_Equip"]) if not pd.isna(row["Cod_Equip"]) else None
                label = f"{int(cod)} - {row.get('DESCRICAO_EQUIPAMENTO','')}" if cod else str(row.get('DESCRICAO_EQUIPAMENTO',''))
                cols = st.columns([3,1,1,1])
                cols[0].markdown(f"**{label}**")
                # mostrar informações principais
                kmc = row.get("Km_Current", np.nan)
                hr_c = row.get("Hr_Current", np.nan)
                cols[1].markdown(f"Km: {kmc if pd.notna(kmc) else '—'}")
                cols[2].markdown(f"Hr: {hr_c if pd.notna(hr_c) else '—'}")
                # checkbox ação
                key = f"manut_done_{cod}"
                if cols[3].checkbox("Manutenção realizada", key=key):
                    # evitar gravação duplicada por sessão
                    if key in st.session_state.manut_processed:
                        st.success("Já registrado nesta sessão.")
                    else:
                        # prepara action dict
                        tipo = "KM" if row.get("Due_Km", False) and not row.get("Due_Hr", False) else ("HR" if row.get("Due_Hr", False) and not row.get("Due_Km", False) else "BOTH")
                        action = {
                            "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Cod_Equip": cod,
                            "DESCRICAO_EQUIPAMENTO": row.get("DESCRICAO_EQUIPAMENTO", ""),
                            "Tipo": tipo,
                            "Km_Current": float(row.get("Km_Current")) if pd.notna(row.get("Km_Current")) else np.nan,
                            "Hr_Current": float(row.get("Hr_Current")) if pd.notna(row.get("Hr_Current")) else np.nan,
                            "Intervalo_KM": int(row.get("Km_Service_Interval")) if pd.notna(row.get("Km_Service_Interval")) else km_interval_default,
                            "Intervalo_HR": int(row.get("Hr_Service_Interval")) if pd.notna(row.get("Hr_Service_Interval")) else hr_interval_default,
                            "Observacao": "",
                            "Usuario": st.session_state.get("user", "usuario_app")
                        }
                        try:
                            append_manut_log(EXCEL_PATH, action)
                            st.success(f"Manutenção registrada no Excel (equip. {cod}).")
                            st.session_state.manut_processed.add(key)
                        except Exception as e:
                            st.error(f"Falha ao registrar manutenção: {e}")

        st.markdown("---")
        st.subheader("Visão geral da frota (manutenção planejada)")
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Km_Current", "Km_Next_Service", "Km_Service_Interval", "Hr_Current", "Hr_Next_Oil", "Hr_Service_Interval"]
        available_over = [c for c in overview_cols if c in mf.columns]
        st.dataframe(mf[available_over].sort_values("Cod_Equip").reset_index(drop=True), use_container_width=True)
        csv_over = mf[available_over].to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar CSV - Plano de Manutenção (Visão Geral)", csv_over, "manutencao_overview.csv", "text/csv")

if __name__ == "__main__":
    main()
