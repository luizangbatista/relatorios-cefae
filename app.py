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
# 🔐 SENHA DE ACESSO
# ==============================

# SENHA_CORRETA = "cefae123"  # ALTERE AQUI

# if "autenticado" not in st.session_state:
#     st.session_state.autenticado = False

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

# ==============================
# 🔐 TELA DE LOGIN
# ==============================

# if not st.session_state.autenticado:
#     topo1, topo2 = st.columns([8, 1])
#     with topo2:
#         st.button(
#             "🌙" if st.session_state.tema == "dark" else "☀️",
#             on_click=alternar_tema,
#             use_container_width=True,
#         )
#
#     st.markdown('<div class="card-dark">', unsafe_allow_html=True)
#     st.markdown(
#         '<div class="home-title">🔒 Acesso restrito</div>',
#         unsafe_allow_html=True,
#     )
#     st.markdown(
#         '<div class="home-subtitle">Digite a senha para acessar o sistema</div>',
#         unsafe_allow_html=True,
#     )
#
#     senha = st.text_input("Senha", type="password")
#
#     if st.button("Entrar", use_container_width=True):
#         if senha == SENHA_CORRETA:
#             st.session_state.autenticado = True
#             st.rerun()
#         else:
#             st.error("Senha incorreta")
#
#     st.markdown("</div>", unsafe_allow_html=True)
#     st.stop()

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


def cadastrar_turma_alunos(nome_turma, texto_alunos):
    turma = nome_turma.strip()
    alunos_lista = [a.strip() for a in texto_alunos.split(";")]
    alunos_lista = [a for a in alunos_lista if a]

    if not turma:
        return False, "Informe o nome da turma."

    if not alunos_lista:
        return False, "Informe pelo menos um aluno separado por ponto e vírgula."

    df_alunos = carregar_alunos()
    df_relatorios = ler_aba("relatorios", COLUNAS_RELATORIOS)

    existentes = set(
        zip(
            df_alunos["turma"].astype(str).str.strip().tolist(),
            df_alunos["aluno"].astype(str).str.strip().tolist(),
        )
    )

    novas_linhas = []
    repetidos = 0

    for aluno in alunos_lista:
        chave = (turma, aluno)
        if chave in existentes:
            repetidos += 1
            continue
        novas_linhas.append({"turma": turma, "aluno": aluno})
        existentes.add(chave)

    if not novas_linhas:
        return False, "Todos os alunos informados já estavam cadastrados nessa turma."

    df_alunos = pd.concat([df_alunos, pd.DataFrame(novas_linhas)], ignore_index=True)
    df_alunos = df_alunos.drop_duplicates().sort_values(["turma", "aluno"]).reset_index(drop=True)
    salvar_abas(df_alunos, df_relatorios)

    return True, f"{len(novas_linhas)} aluno(s) cadastrado(s) na turma '{turma}'. Repetidos ignorados: {repetidos}."


def salvar_relatorio(data_relatorio, turma, monitor, alunos, texto_relatorio):
    df_alunos = ler_aba("alunos", COLUNAS_ALUNOS)
    df_relatorios = carregar_relatorios()

    turma = str(turma).strip()
    monitor = str(monitor).strip()
    texto_relatorio = str(texto_relatorio).strip()
    alunos_texto = "; ".join(alunos)

    if not turma:
        return False, "Selecione uma turma."
    if not monitor:
        return False, "Selecione um monitor."
    if not alunos:
        return False, "Selecione pelo menos um aluno."
    if not texto_relatorio:
        return False, "Escreva o relatório."

    nova_linha = pd.DataFrame(
        [{
            "data": data_relatorio.strftime("%Y-%m-%d"),
            "turma": turma,
            "monitor": monitor,
            "alunos": alunos_texto,
            "relatorio": texto_relatorio,
        }]
    )

    df_relatorios_base = (
        df_relatorios[COLUNAS_RELATORIOS].copy()
        if not df_relatorios.empty
        else pd.DataFrame(columns=COLUNAS_RELATORIOS)
    )
    df_relatorios_base = pd.concat([df_relatorios_base, nova_linha], ignore_index=True)

    salvar_abas(df_alunos, df_relatorios_base)

    return True, "Relatório salvo com sucesso."


def filtrar_relatorios(df, turma=None, aluno=None, monitor=None, data_ini=None, data_fim=None):
    if df.empty:
        return df.copy()

    filtrado = df.copy()

    if turma and turma != "Todas":
        filtrado = filtrado[filtrado["turma"] == turma]

    if monitor and monitor != "Todos":
        filtrado = filtrado[filtrado["monitor"] == monitor]

    if aluno and aluno != "Todos":
        filtrado = filtrado[
            filtrado["alunos"].apply(
                lambda x: aluno in [parte.strip() for parte in str(x).split(";") if parte.strip()]
            )
        ]

    if data_ini:
        filtrado = filtrado[filtrado["data_dt"] >= pd.Timestamp(data_ini)]

    if data_fim:
        filtrado = filtrado[filtrado["data_dt"] <= pd.Timestamp(data_fim)]

    filtrado = filtrado.sort_values(["data_dt"], ascending=[False]).reset_index(drop=True)
    return filtrado


def gerar_texto_filtros_utilizados(turma_filtro, aluno_filtro, monitor_filtro, data_ini, data_fim):
    filtros = []

    if turma_filtro and turma_filtro != "Todas":
        filtros.append(f"Turma: {turma_filtro}")

    if aluno_filtro and aluno_filtro != "Todos":
        filtros.append(f"Aluno: {aluno_filtro}")

    if monitor_filtro and monitor_filtro != "Todos":
        filtros.append(f"Monitor: {monitor_filtro}")

    if data_ini:
        filtros.append(f"Data inicial: {data_ini.strftime('%d/%m/%Y')}")

    if data_fim:
        filtros.append(f"Data final: {data_fim.strftime('%d/%m/%Y')}")

    if not filtros:
        return ""

    return " | ".join(filtros)


def gerar_pdf_relatorios(df, filtros_texto):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    largura, altura = A4

    margem_esq = 1.5 * 28.35
    margem_dir = 1.5 * 28.35
    margem_topo = 4.5 * 28.35
    margem_base = 4.0 * 28.35

    largura_texto = largura - margem_esq - margem_dir
    y = altura - margem_topo

    fonte_normal = "Helvetica"
    fonte_negrito = "Helvetica-Bold"
    tamanho = 11
    espacamento_linha = 16.5
    espacamento_relatorio = 28.35

    def desenhar_timbrado():
        if os.path.exists(ARQUIVO_TIMBRADO):
            c.drawImage(
                ARQUIVO_TIMBRADO,
                0,
                0,
                width=largura,
                height=altura
            )

    def nova_pagina():
        nonlocal y
        c.showPage()
        desenhar_timbrado()
        y = altura - margem_topo

    def escrever_linha_centralizada(texto, fonte=fonte_normal, tamanho_fonte=tamanho, espaco=18):
        nonlocal y
        linhas = simpleSplit(str(texto), fonte, tamanho_fonte, largura_texto)
        c.setFont(fonte, tamanho_fonte)

        for linha in linhas:
            if y < margem_base:
                nova_pagina()
                c.setFont(fonte, tamanho_fonte)
            largura_linha = c.stringWidth(linha, fonte, tamanho_fonte)
            x = margem_esq + (largura_texto - largura_linha) / 2
            c.drawString(x, y, linha)
            y -= espaco

    def escrever_texto_justificado(rotulo, texto, espaco_linha=espacamento_linha):
        nonlocal y

        rotulo = str(rotulo)
        texto = str(texto).strip()

        c.setFont(fonte_negrito, tamanho)
        largura_rotulo = c.stringWidth(rotulo, fonte_negrito, tamanho)
        largura_primeira = largura_texto - largura_rotulo

        palavras = texto.split()
        linhas = []
        linha_atual = ""
        largura_limite = largura_primeira

        c.setFont(fonte_normal, tamanho)

        for palavra in palavras:
            teste = palavra if not linha_atual else f"{linha_atual} {palavra}"
            if c.stringWidth(teste, fonte_normal, tamanho) <= largura_limite:
                linha_atual = teste
            else:
                if linha_atual:
                    linhas.append(linha_atual)
                linha_atual = palavra
                largura_limite = largura_texto

        if linha_atual:
            linhas.append(linha_atual)

        if not linhas:
            linhas = [""]

        for i, linha in enumerate(linhas):
            if y < margem_base:
                nova_pagina()

            if i == 0:
                c.setFont(fonte_negrito, tamanho)
                c.drawString(margem_esq, y, rotulo)
                c.setFont(fonte_normal, tamanho)
                c.drawString(margem_esq + largura_rotulo, y, linha)
            else:
                palavras_linha = linha.split()
                if len(palavras_linha) > 1 and i != len(linhas) - 1:
                    largura_palavras = sum(c.stringWidth(p, fonte_normal, tamanho) for p in palavras_linha)
                    espaco_total = largura_texto - largura_palavras
                    espaco = espaco_total / (len(palavras_linha) - 1)
                    x = margem_esq
                    c.setFont(fonte_normal, tamanho)
                    for palavra in palavras_linha[:-1]:
                        c.drawString(x, y, palavra)
                        x += c.stringWidth(palavra, fonte_normal, tamanho) + espaco
                    c.drawString(x, y, palavras_linha[-1])
                else:
                    c.setFont(fonte_normal, tamanho)
                    c.drawString(margem_esq, y, linha)

            y -= espaco_linha

    def escrever_linha_mista(data_str, monitor, turma, espaco=espacamento_linha):
        nonlocal y

        partes = [
            ("Data:", True),
            (f" {data_str} | ", False),
            ("Monitor:", True),
            (f" {monitor} | ", False),
            ("Turma:", True),
            (f" {turma}", False),
        ]

        linhas = []
        linha_atual = []
        largura_atual = 0

        for texto, negrito in partes:
            fonte = fonte_negrito if negrito else fonte_normal
            largura_parte = c.stringWidth(texto, fonte, tamanho)

            if largura_atual + largura_parte <= largura_texto:
                linha_atual.append((texto, negrito))
                largura_atual += largura_parte
            else:
                if linha_atual:
                    linhas.append(linha_atual)
                linha_atual = [(texto, negrito)]
                largura_atual = largura_parte

        if linha_atual:
            linhas.append(linha_atual)

        for linha in linhas:
            if y < margem_base:
                nova_pagina()

            x = margem_esq
            for texto, negrito in linha:
                fonte = fonte_negrito if negrito else fonte_normal
                c.setFont(fonte, tamanho)
                c.drawString(x, y, texto)
                x += c.stringWidth(texto, fonte, tamanho)

            y -= espaco

    desenhar_timbrado()

    if filtros_texto:
        escrever_linha_centralizada(filtros_texto, fonte_normal, tamanho, 18)
        y -= 10

    if df.empty:
        c.setFont(fonte_negrito, tamanho)
        c.drawString(margem_esq, y, "Nenhum relatório encontrado.")
    else:
        for i, (_, row) in enumerate(df.iterrows()):
            try:
                data_formatada = pd.to_datetime(row["data"]).strftime("%d/%m")
            except Exception:
                data_formatada = str(row["data"])

            if i > 0:
                y -= espacamento_relatorio

            escrever_linha_mista(
                data_formatada,
                str(row.get("monitor", "")),
                str(row.get("turma", "")),
            )
            escrever_texto_justificado("Alunos: ", str(row.get("alunos", "")))
            escrever_texto_justificado("Relatório: ", str(row.get("relatorio", "")))

    c.save()
    buffer.seek(0)
    return buffer


def gerar_docx_relatorios(df, filtros_texto):
    doc = Document()

    sec = doc.sections[0]
    sec.top_margin = Pt(127.56)
    sec.bottom_margin = Pt(113.4)
    sec.left_margin = Pt(42.52)
    sec.right_margin = Pt(42.52)

    estilo_normal = doc.styles["Normal"]
    estilo_normal.font.name = "Calibri"
    estilo_normal.font.size = Pt(11)

    if filtros_texto:
        p_filtros = doc.add_paragraph()
        p_filtros.alignment = 1
        p_filtros.paragraph_format.line_spacing = 1.5
        r_filtros = p_filtros.add_run(filtros_texto)
        r_filtros.font.name = "Calibri"
        r_filtros.font.size = Pt(11)

    if df.empty:
        p = doc.add_paragraph()
        r = p.add_run("Nenhum relatório encontrado.")
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(11)
    else:
        for i, (_, row) in enumerate(df.iterrows()):
            try:
                data_formatada = pd.to_datetime(row.get("data", ""), errors="coerce").strftime("%d/%m")
            except Exception:
                data_formatada = str(row.get("data", ""))

            if i > 0:
                p_espaco = doc.add_paragraph()
                p_espaco.paragraph_format.space_before = Pt(28.35)

            p1 = doc.add_paragraph()
            p1.paragraph_format.line_spacing = 1.5

            r = p1.add_run("Data: ")
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            r = p1.add_run(f"{data_formatada} | ")
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            r = p1.add_run("Monitor: ")
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            r = p1.add_run(f"{row.get('monitor', '')} | ")
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            r = p1.add_run("Turma: ")
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            r = p1.add_run(str(row.get("turma", "")))
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            p2 = doc.add_paragraph()
            p2.paragraph_format.line_spacing = 1.5
            r = p2.add_run("Alunos: ")
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)
            r = p2.add_run(str(row.get("alunos", "")))
            r.font.name = "Calibri"
            r.font.size = Pt(11)

            p3 = doc.add_paragraph()
            p3.paragraph_format.line_spacing = 1.5
            p3.paragraph_format.alignment = 3
            r = p3.add_run("Relatório: ")
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(11)
            r = p3.add_run(str(row.get("relatorio", "")))
            r.font.name = "Calibri"
            r.font.size = Pt(11)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def deletar_relatorios(df_filtrado, indices_filtrados):
    if not indices_filtrados:
        return False, "Selecione pelo menos um relatório para excluir."

    df_alunos = ler_aba("alunos", COLUNAS_ALUNOS)
    df_relatorios_completo = carregar_relatorios()

    linhas_para_remover = df_filtrado.loc[indices_filtrados, COLUNAS_RELATORIOS].copy()
    restantes = df_relatorios_completo.copy()

    for _, linha in linhas_para_remover.iterrows():
        mascara = (
            (restantes["data"] == linha["data"]) &
            (restantes["turma"] == linha["turma"]) &
            (restantes["monitor"] == linha["monitor"]) &
            (restantes["alunos"] == linha["alunos"]) &
            (restantes["relatorio"] == linha["relatorio"])
        )
        idx_match = restantes[mascara].index
        if len(idx_match) > 0:
            restantes = restantes.drop(idx_match[0])

    restantes = restantes[COLUNAS_RELATORIOS].reset_index(drop=True)
    salvar_abas(df_alunos, restantes)

    return True, f"{len(indices_filtrados)} relatório(s) excluído(s) com sucesso."


def deletar_turma(nome_turma, excluir_relatorios=False):
    df_alunos = ler_aba("alunos", COLUNAS_ALUNOS)
    df_relatorios = ler_aba("relatorios", COLUNAS_RELATORIOS)

    nome_turma = str(nome_turma).strip()

    if not nome_turma:
        return False, "Selecione uma turma."

    if df_alunos.empty:
        return False, "Não há turmas cadastradas."

    turmas_existentes = df_alunos["turma"].astype(str).str.strip().unique()
    if nome_turma not in turmas_existentes:
        return False, "Turma não encontrada."

    df_alunos = df_alunos[
        df_alunos["turma"].astype(str).str.strip() != nome_turma
    ].reset_index(drop=True)

    if excluir_relatorios:
        df_relatorios = df_relatorios[
            df_relatorios["turma"].astype(str).str.strip() != nome_turma
        ].reset_index(drop=True)

    salvar_abas(df_alunos, df_relatorios)

    if excluir_relatorios:
        return True, f"Turma '{nome_turma}' e seus relatórios foram excluídos com sucesso."
    else:
        return True, f"Turma '{nome_turma}' foi excluída com sucesso. Os relatórios antigos foram mantidos."


def ir_para(nome_pagina):
    st.session_state.pagina = nome_pagina
    st.rerun()


def sair():
    # st.session_state.autenticado = False
    st.session_state.pagina = "home"
    st.session_state.modo_exclusao = False
    st.rerun()


def topo_app():
    c1, c2, c3 = st.columns([7, 1, 1])
    with c2:
        st.button(
            "🌙" if st.session_state.tema == "dark" else "☀️",
            on_click=alternar_tema,
            use_container_width=True,
        )
    with c3:
        st.button("Sair", on_click=sair, use_container_width=True)


def botao_voltar():
    if st.button("⬅️ Voltar para a página inicial"):
        st.session_state.modo_exclusao = False
        ir_para("home")


def tela_home():
    topo_app()

    st.markdown('<div class="home-title">📚 Sistema de Monitoria</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="home-subtitle">Selecione uma das opções abaixo</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.mensagem_sucesso:
        st.markdown(
            f'<div class="success-box">{st.session_state.mensagem_sucesso}</div>',
            unsafe_allow_html=True,
        )
        st.session_state.mensagem_sucesso = ""

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("Cadastrar turma", use_container_width=True):
            st.session_state.modo_exclusao = False
            ir_para("cadastrar_turma")

    with c2:
        if st.button("Enviar novo relatório", use_container_width=True):
            st.session_state.modo_exclusao = False
            ir_para("cadastrar_relatorio")

    with c3:
        if st.button("Consultar relatórios enviados", use_container_width=True):
            st.session_state.modo_exclusao = False
            ir_para("consultar")


inicializar_arquivo()

df_alunos = carregar_alunos()
df_relatorios = carregar_relatorios()
pagina = st.session_state.pagina

if pagina == "home":
    tela_home()

elif pagina == "cadastrar_turma":
    topo_app()
    botao_voltar()
    st.title("Cadastrar turma")

    st.markdown('<div class="card-dark">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Cadastro de turma e alunos</div>', unsafe_allow_html=True)

    with st.form("form_turma"):
        nome_turma = st.text_input("Nome da turma", placeholder="Ex.: Sexto A")
        texto_alunos = st.text_area(
            "Alunos separados por ponto e vírgula",
            height=180,
            placeholder="Ex.: Ana Souza; Bruno Lima; Carla Mendes",
        )
        enviar_turma = st.form_submit_button("Salvar turma e alunos")

    st.markdown("</div>", unsafe_allow_html=True)

    if enviar_turma:
        ok, mensagem = cadastrar_turma_alunos(nome_turma, texto_alunos)
        if ok:
            st.success(mensagem)
            st.rerun()
        else:
            st.error(mensagem)

    st.markdown('<div class="card-dark">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Turmas cadastradas</div>', unsafe_allow_html=True)

    if df_alunos.empty:
        st.info("Nenhuma turma cadastrada ainda.")
    else:
        resumo = (
            df_alunos.groupby("turma", as_index=False)
            .agg(total_alunos=("aluno", "count"))
            .sort_values("turma")
        )
        st.dataframe(resumo, use_container_width=True, hide_index=True)

        turma_visualizar = st.selectbox(
            "Visualizar alunos da turma",
            options=sorted(df_alunos["turma"].unique().tolist()),
        )
        alunos_turma = df_alunos[df_alunos["turma"] == turma_visualizar]["aluno"].tolist()

        st.write("**Alunos cadastrados nessa turma:**")
        for aluno in alunos_turma:
            st.write(f"- {aluno}")

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card-dark">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Excluir turma</div>', unsafe_allow_html=True)

    turmas_existentes = sorted(df_alunos["turma"].unique().tolist()) if not df_alunos.empty else []

    if not turmas_existentes:
        st.info("Nenhuma turma disponível para excluir.")
    else:
        turma_excluir = st.selectbox(
            "Selecione a turma para excluir",
            options=turmas_existentes,
            key="turma_excluir"
        )

        excluir_relatorios_tambem = st.checkbox(
            "Excluir também os relatórios dessa turma",
            key="check_excluir_relatorios_turma"
        )

        if st.button("Excluir turma", use_container_width=True):
            ok, mensagem = deletar_turma(turma_excluir, excluir_relatorios_tambem)
            if ok:
                st.success(mensagem)
                st.rerun()
            else:
                st.error(mensagem)

    st.markdown("</div>", unsafe_allow_html=True)

elif pagina == "cadastrar_relatorio":
    topo_app()
    botao_voltar()
    st.title("Enviar novo relatório")

    turmas_disponiveis = sorted(df_alunos["turma"].unique().tolist()) if not df_alunos.empty else []

    if not turmas_disponiveis:
        st.warning("Cadastre pelo menos uma turma antes de criar relatórios.")
    else:
        st.markdown('<div class="card-dark">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Cadastro de relatório</div>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)

        with col1:
            data_relatorio = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")

        with col2:
            turma_escolhida = st.selectbox("Turma", options=turmas_disponiveis)

        with col3:
            monitor_escolhido = st.selectbox("Monitor", options=MONITORES)

        alunos_da_turma = (
            df_alunos[df_alunos["turma"] == turma_escolhida]["aluno"]
            .dropna()
            .astype(str)
            .sort_values()
            .tolist()
        )

        st.markdown("### Seleção de alunos")

        chave_alunos = f"alunos_selecionados_{turma_escolhida}"
        if chave_alunos not in st.session_state:
            st.session_state[chave_alunos] = []

        b1, b2 = st.columns(2)

        with b1:
            if st.button("Selecionar todos"):
                st.session_state[chave_alunos] = alunos_da_turma.copy()
                st.rerun()

        with b2:
            if st.button("Selecionar nenhum"):
                st.session_state[chave_alunos] = []
                st.rerun()

        alunos_selecionados = st.multiselect(
            "Alunos da turma",
            options=alunos_da_turma,
            default=st.session_state[chave_alunos],
        )
        st.session_state[chave_alunos] = alunos_selecionados

        texto_relatorio = st.text_area(
            "Relatório",
            height=220,
            placeholder="Escreva aqui o relatório da monitoria...",
        )

        if st.button("Salvar relatório"):
            ok, mensagem = salvar_relatorio(
                data_relatorio=data_relatorio,
                turma=turma_escolhida,
                monitor=monitor_escolhido,
                alunos=alunos_selecionados,
                texto_relatorio=texto_relatorio,
            )
            if ok:
                st.session_state[chave_alunos] = []
                st.session_state.mensagem_sucesso = mensagem
                st.session_state.modo_exclusao = False
                ir_para("home")
            else:
                st.error(mensagem)

        st.markdown("</div>", unsafe_allow_html=True)

elif pagina == "consultar":
    topo_app()
    botao_voltar()
    st.title("Consultar relatórios enviados")

    if df_relatorios.empty:
        st.info("Nenhum relatório cadastrado ainda.")
    else:
        st.markdown('<div class="card-dark">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)

        turmas = ["Todas"] + sorted(df_relatorios["turma"].dropna().astype(str).unique().tolist())
        monitores = ["Todos"] + sorted(df_relatorios["monitor"].dropna().astype(str).unique().tolist())

        todos_alunos = set()
        for texto in df_relatorios["alunos"].fillna("").astype(str):
            for parte in texto.split(";"):
                nome = parte.strip()
                if nome:
                    todos_alunos.add(nome)
        alunos_filtro = ["Todos"] + sorted(todos_alunos)

        c1, c2, c3 = st.columns(3)
        with c1:
            turma_filtro = st.selectbox("Filtrar por turma", options=turmas)
        with c2:
            aluno_filtro = st.selectbox("Filtrar por aluno", options=alunos_filtro)
        with c3:
            monitor_filtro = st.selectbox("Filtrar por monitor", options=monitores)

        usar_filtro_data = st.checkbox("Filtrar por data")

        if usar_filtro_data:
            c4, c5 = st.columns(2)
            with c4:
                data_ini = st.date_input("Data inicial", value=date.today(), format="DD/MM/YYYY", key="data_ini_consulta")
            with c5:
                data_fim = st.date_input("Data final", value=date.today(), format="DD/MM/YYYY", key="data_fim_consulta")
        else:
            data_ini = None
            data_fim = None

        st.markdown("</div>", unsafe_allow_html=True)

        df_filtrado = filtrar_relatorios(
            df=df_relatorios,
            turma=turma_filtro,
            aluno=aluno_filtro,
            monitor=monitor_filtro,
            data_ini=data_ini,
            data_fim=data_fim,
        )

        filtros_texto = gerar_texto_filtros_utilizados(
            turma_filtro, aluno_filtro, monitor_filtro, data_ini, data_fim
        )

        st.markdown('<div class="card-dark">', unsafe_allow_html=True)
        st.markdown(f"**Total encontrado:** {len(df_filtrado)} relatório(s)")

        if df_filtrado.empty:
            st.warning("Nenhum relatório corresponde aos filtros selecionados.")
            st.session_state.modo_exclusao = False
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            c_download1, c_download2 = st.columns(2)

            with c_download1:
                pdf_bytes = gerar_pdf_relatorios(df_filtrado, filtros_texto)
                st.download_button(
                    label="Gerar arquivo PDF",
                    data=pdf_bytes,
                    file_name=f"relatorio_cefae_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            with c_download2:
                docx_bytes = gerar_docx_relatorios(df_filtrado, filtros_texto)
                st.download_button(
                    label="Gerar arquivo DOC",
                    data=docx_bytes,
                    file_name=f"relatorio_cefae_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            c_botao1, c_botao2 = st.columns(2)

            with c_botao1:
                if not st.session_state.modo_exclusao:
                    if st.button("Excluir relatórios", use_container_width=True):
                        st.session_state.modo_exclusao = True
                        st.rerun()

            with c_botao2:
                if st.session_state.modo_exclusao:
                    if st.button("Cancelar exclusão", use_container_width=True):
                        st.session_state.modo_exclusao = False
                        st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.modo_exclusao:
                st.markdown('<div class="card-soft">', unsafe_allow_html=True)
                st.markdown("### Selecione os relatórios para excluir")
                st.markdown("</div>", unsafe_allow_html=True)

            indices_para_excluir = []

            for idx, row in df_filtrado.iterrows():
                try:
                    data_formatada = pd.to_datetime(row["data"]).strftime("%d/%m/%Y")
                except Exception:
                    data_formatada = str(row["data"])

                st.markdown('<div class="card-dark">', unsafe_allow_html=True)

                if st.session_state.modo_exclusao:
                    marcado = st.checkbox(
                        "Selecionar para excluir",
                        key=f"excluir_relatorio_{idx}"
                    )
                    if marcado:
                        indices_para_excluir.append(idx)

                st.write(f"**Data:** {data_formatada}")
                st.write(f"**Turma:** {row['turma']}")
                st.write(f"**Monitor:** {row['monitor']}")
                st.write(f"**Alunos:** {row['alunos']}")
                st.write(f"**Relatório:** {row['relatorio']}")

                st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.modo_exclusao:
                if st.button("Excluir selecionados", use_container_width=True):
                    ok, mensagem = deletar_relatorios(df_filtrado, indices_para_excluir)
                    if ok:
                        for idx in df_filtrado.index:
                            chave = f"excluir_relatorio_{idx}"
                            if chave in st.session_state:
                                del st.session_state[chave]
                        st.session_state.modo_exclusao = False
                        st.session_state.mensagem_sucesso = mensagem
                        ir_para("consultar")
                    else:
                        st.error(mensagem)