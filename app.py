import io
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

st.set_page_config(
    page_title="Relatórios de Monitoria",
    page_icon="📝",
    layout="wide",
)

# ==============================
# 🌗 TEMA
# ==============================

if "tema" not in st.session_state:
    st.session_state.tema = "light"


def alternar_tema():
    st.session_state.tema = "light" if st.session_state.tema == "dark" else "dark"


if st.session_state.tema == "dark":
    BG = "#0b1220"
    CARD = "#111827"
    CARD_SOFT = "#0f172a"
    BORDER = "#243041"
    TEXT = "#f8fafc"
    SUBTEXT = "#cbd5e1"
    SUCCESS_BG = "#052e1a"
    SUCCESS_BORDER = "#166534"
    SUCCESS_TEXT = "#bbf7d0"
    BUTTON_HOVER = "#172033"
else:
    BG = "#f8fafc"
    CARD = "#ffffff"
    CARD_SOFT = "#f1f5f9"
    BORDER = "#dbe4ee"
    TEXT = "#111827"
    SUBTEXT = "#4b5563"
    SUCCESS_BG = "#ecfdf3"
    SUCCESS_BORDER = "#86efac"
    SUCCESS_TEXT = "#166534"
    BUTTON_HOVER = "#eef2f7"

st.markdown(
    f"""
    <style>
    .block-container {{
        padding-top: 1.1rem;
        padding-bottom: 2rem;
        max-width: 920px;
    }}

    html, body, [data-testid="stAppViewContainer"] {{
        background-color: {BG};
    }}

    [data-testid="stHeader"] {{
        background: transparent;
    }}

    [data-testid="stToolbar"] {{
        right: 0.5rem;
    }}

    div[data-testid="stForm"] {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 16px;
        padding: 1rem;
    }}

    .card-dark {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 16px;
        padding: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 0 rgba(0,0,0,0.02);
    }}

    .card-soft {{
        background: {CARD_SOFT};
        border: 1px solid {BORDER};
        border-radius: 16px;
        padding: 0.9rem;
        margin-bottom: 1rem;
    }}

    .home-title {{
        text-align: center;
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
        color: {TEXT};
    }}

    .home-subtitle {{
        text-align: center;
        font-size: 1rem;
        color: {SUBTEXT};
        margin-bottom: 1.4rem;
    }}

    .success-box {{
        padding: 0.9rem 1rem;
        border-radius: 14px;
        background: {SUCCESS_BG};
        border: 1px solid {SUCCESS_BORDER};
        color: {SUCCESS_TEXT};
        margin-bottom: 1rem;
        text-align: center;
        font-weight: 600;
    }}

    .section-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: {TEXT};
        margin-bottom: 0.6rem;
    }}

    .stButton > button,
    .stDownloadButton > button {{
        width: 100%;
        border-radius: 14px;
        min-height: 52px;
        font-size: 16px;
        background: {CARD};
        color: {TEXT};
        border: 1px solid {BORDER};
    }}

    .stButton > button:hover,
    .stDownloadButton > button:hover {{
        border-color: #60a5fa;
        color: {TEXT};
        background: {BUTTON_HOVER};
    }}

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-baseweb="textarea"] > div {{
        background-color: {CARD} !important;
        border-color: {BORDER} !important;
        color: {TEXT} !important;
        border-radius: 12px !important;
    }}

    input, textarea {{
        color: {TEXT} !important;
    }}

    label, .stMarkdown, .stText, p, span, div {{
        color: {TEXT};
    }}

    div[data-testid="stDateInput"] > div {{
        background-color: {CARD} !important;
        border-radius: 12px !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

ARQUIVO_DADOS = "dados_monitoria.xlsx"
ARQUIVO_TIMBRADO = "timbrado.png"

MONITORES = [
    "Arthur - Matemática",
    "Davi - Ciências",
    "Dayane - História",
    "Gabriel - Física",
    "Gabriel - Português",
    "Lorraine - 4º ano",
    "Luiza - Matemática",
    "Maria Eduarda - 5º ano",
    "Raphael - Matemática",
    "Rayanne - 5º ano",
    "Roberta - 4º ano",
    "Silvana - Coordenação",
    "Uill - Português",
    "Vinícius - Inglês"
]

COLUNAS_ALUNOS = ["turma", "aluno"]
COLUNAS_RELATORIOS = ["data", "turma", "monitor", "alunos", "relatorio"]

if "pagina" not in st.session_state:
    st.session_state.pagina = "home"

if "mensagem_sucesso" not in st.session_state:
    st.session_state.mensagem_sucesso = ""

if "modo_exclusao" not in st.session_state:
    st.session_state.modo_exclusao = False


def inicializar_arquivo():
    if not os.path.exists(ARQUIVO_DADOS):
        with pd.ExcelWriter(ARQUIVO_DADOS, engine="openpyxl") as writer:
            pd.DataFrame(columns=COLUNAS_ALUNOS).to_excel(
                writer, sheet_name="alunos", index=False
            )
            pd.DataFrame(columns=COLUNAS_RELATORIOS).to_excel(
                writer, sheet_name="relatorios", index=False
            )


def ler_aba(nome_aba, colunas):
    inicializar_arquivo()
    try:
        df = pd.read_excel(ARQUIVO_DADOS, sheet_name=nome_aba, engine="openpyxl")
    except Exception:
        df = pd.DataFrame(columns=colunas)

    for col in colunas:
        if col not in df.columns:
            df[col] = ""

    return df[colunas].copy()


def salvar_abas(df_alunos, df_relatorios):
    with pd.ExcelWriter(ARQUIVO_DADOS, engine="openpyxl", mode="w") as writer:
        df_alunos.to_excel(writer, sheet_name="alunos", index=False)
        df_relatorios.to_excel(writer, sheet_name="relatorios", index=False)


def carregar_alunos():
    df = ler_aba("alunos", COLUNAS_ALUNOS)
    if not df.empty:
        df["turma"] = df["turma"].astype(str).str.strip()
        df["aluno"] = df["aluno"].astype(str).str.strip()
        df = df[(df["turma"] != "") & (df["aluno"] != "")]
        df = df.drop_duplicates().sort_values(["turma", "aluno"]).reset_index(drop=True)
    return df


def carregar_relatorios():
    df = ler_aba("relatorios", COLUNAS_RELATORIOS)
    if not df.empty:
        for col in COLUNAS_RELATORIOS:
            df[col] = df[col].astype(str).fillna("").str.strip()
        df["data_dt"] = pd.to_datetime(df["data"], errors="coerce")
        df = df.sort_values(["data_dt"], ascending=[False]).reset_index(drop=True)
    return df

# ... (continua exatamente igual ao seu código original)
