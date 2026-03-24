import io
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfgen import canvas

st.set_page_config(
    page_title="Relatórios de Monitoria",
    page_icon="📝",
    layout="wide",
)

ARQUIVO_DADOS = "dados_monitoria.xlsx"

MONITORES = [
    "Luiza Matemática",
    "Gabriel Português",
]

COLUNAS_ALUNOS = ["turma", "aluno"]
COLUNAS_RELATORIOS = ["id", "data", "turma", "monitor", "alunos", "relatorio"]


st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    .stButton > button {
        width: 100%;
    }
    .caixa {
        padding: 1rem;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


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
    with pd.ExcelWriter(
        ARQUIVO_DADOS,
        engine="openpyxl",
        mode="w",
    ) as writer:
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
        df["id_num"] = pd.to_numeric(df["id"], errors="coerce")
        df["data_dt"] = pd.to_datetime(df["data"], errors="coerce")
        df = df.sort_values(["data_dt", "id_num"], ascending=[False, False]).reset_index(drop=True)
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


def proximo_id_relatorio(df_relatorios):
    if df_relatorios.empty:
        return 1
    ids = pd.to_numeric(df_relatorios["id"], errors="coerce").dropna()
    if ids.empty:
        return 1
    return int(ids.max()) + 1


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

    novo_id = proximo_id_relatorio(df_relatorios)

    nova_linha = pd.DataFrame(
        [{
            "id": novo_id,
            "data": data_relatorio.strftime("%Y-%m-%d"),
            "turma": turma,
            "monitor": monitor,
            "alunos": alunos_texto,
            "relatorio": texto_relatorio,
        }]
    )

    df_relatorios = pd.concat([df_relatorios[COLUNAS_RELATORIOS], nova_linha], ignore_index=True)
    salvar_abas(df_alunos, df_relatorios[COLUNAS_RELATORIOS])

    return True, f"Relatório {novo_id} salvo com sucesso."


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

    filtrado = filtrado.sort_values(["data_dt", "id_num"], ascending=[False, False]).reset_index(drop=True)
    return filtrado


def gerar_pdf_relatorios(df, filtros_texto):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    largura, altura = A4
    margem_esq = 40
    margem_dir = 40
    y = altura - 40
    largura_texto = largura - margem_esq - margem_dir

    def nova_pagina():
        nonlocal y
        c.showPage()
        y = altura - 40

    def escrever_linha(texto, fonte="Helvetica", tamanho=10, espaco=14):
        nonlocal y
        linhas = simpleSplit(str(texto), fonte, tamanho, largura_texto)
        c.setFont(fonte, tamanho)
        for linha in linhas:
            if y < 50:
                nova_pagina()
                c.setFont(fonte, tamanho)
            c.drawString(margem_esq, y, linha)
            y -= espaco

    escrever_linha("Relatórios de Monitoria", "Helvetica-Bold", 16, 18)
    escrever_linha(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "Helvetica", 10, 14)
    escrever_linha(f"Filtros aplicados: {filtros_texto}", "Helvetica", 10, 16)
    y -= 4

    if df.empty:
        escrever_linha("Nenhum relatório encontrado.", "Helvetica-Bold", 11, 16)
    else:
        for _, row in df.iterrows():
            try:
                data_formatada = pd.to_datetime(row.get("data", ""), errors="coerce").strftime("%d/%m/%Y")
            except Exception:
                data_formatada = str(row.get("data", ""))

            escrever_linha("-" * 90, "Helvetica", 9, 12)
            escrever_linha(f"ID: {row.get('id', '')}", "Helvetica-Bold", 10, 14)
            escrever_linha(f"Data: {data_formatada}", "Helvetica", 10, 14)
            escrever_linha(f"Turma: {row.get('turma', '')}", "Helvetica", 10, 14)
            escrever_linha(f"Monitor: {row.get('monitor', '')}", "Helvetica", 10, 14)
            escrever_linha(f"Alunos: {row.get('alunos', '')}", "Helvetica", 10, 14)
            escrever_linha("Relatório:", "Helvetica-Bold", 10, 14)
            escrever_linha(f"{row.get('relatorio', '')}", "Helvetica", 10, 15)
            y -= 8

            if y < 80:
                nova_pagina()

    c.save()
    buffer.seek(0)
    return buffer


inicializar_arquivo()

st.title("📝 Relatórios de Monitoria")

menu = st.sidebar.radio(
    "Navegação",
    ["Cadastrar turma", "Cadastrar relatório", "Consultar relatórios"],
)

df_alunos = carregar_alunos()
df_relatorios = carregar_relatorios()

if menu == "Cadastrar turma":
    st.subheader("Cadastro de turma e alunos")

    with st.form("form_turma"):
        nome_turma = st.text_input("Nome da turma", placeholder="Ex.: Sexto A")
        texto_alunos = st.text_area(
            "Alunos separados por ponto e vírgula",
            height=180,
            placeholder="Ex.: Ana Souza; Bruno Lima; Carla Mendes",
        )
        enviar_turma = st.form_submit_button("Salvar turma e alunos")

    if enviar_turma:
        ok, mensagem = cadastrar_turma_alunos(nome_turma, texto_alunos)
        if ok:
            st.success(mensagem)
            st.rerun()
        else:
            st.error(mensagem)

    st.markdown("---")
    st.subheader("Turmas cadastradas")

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
        st.write("**Alunos:**")
        for aluno in alunos_turma:
            st.write(f"- {aluno}")

elif menu == "Cadastrar relatório":
    st.subheader("Cadastro de relatório")

    turmas_disponiveis = sorted(df_alunos["turma"].unique().tolist()) if not df_alunos.empty else []

    if not turmas_disponiveis:
        st.warning("Cadastre pelo menos uma turma antes de criar relatórios.")
    else:
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
        if b1.button("Selecionar todos"):
            st.session_state[chave_alunos] = alunos_da_turma.copy()
            st.rerun()

        if b2.button("Selecionar nenhum"):
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
                st.success(mensagem)
                st.rerun()
            else:
                st.error(mensagem)

elif menu == "Consultar relatórios":
    st.subheader("Consulta de relatórios")

    if df_relatorios.empty:
        st.info("Nenhum relatório cadastrado ainda.")
    else:
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

        c4, c5 = st.columns(2)
        with c4:
            data_ini = st.date_input("Data inicial", value=None, format="DD/MM/YYYY")
        with c5:
            data_fim = st.date_input("Data final", value=None, format="DD/MM/YYYY")

        df_filtrado = filtrar_relatorios(
            df=df_relatorios,
            turma=turma_filtro,
            aluno=aluno_filtro,
            monitor=monitor_filtro,
            data_ini=data_ini,
            data_fim=data_fim,
        )

        st.markdown(f"**Total encontrado:** {len(df_filtrado)} relatório(s)")

        if df_filtrado.empty:
            st.warning("Nenhum relatório corresponde aos filtros selecionados.")
        else:
            filtros_texto = (
                f"Turma: {turma_filtro} | "
                f"Aluno: {aluno_filtro} | "
                f"Monitor: {monitor_filtro} | "
                f"Data inicial: {data_ini.strftime('%d/%m/%Y') if data_ini else '-'} | "
                f"Data final: {data_fim.strftime('%d/%m/%Y') if data_fim else '-'}"
            )

            pdf_bytes = gerar_pdf_relatorios(df_filtrado, filtros_texto)

            st.download_button(
                label="Baixar PDF dos relatórios filtrados",
                data=pdf_bytes,
                file_name=f"relatorios_monitoria_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
            )

            st.markdown("---")

            for _, row in df_filtrado.iterrows():
                try:
                    data_formatada = pd.to_datetime(row["data"]).strftime("%d/%m/%Y")
                except Exception:
                    data_formatada = str(row["data"])

                with st.container():
                    st.markdown('<div class="caixa">', unsafe_allow_html=True)
                    st.write(f"**ID:** {row['id']}")
                    st.write(f"**Data:** {data_formatada}")
                    st.write(f"**Turma:** {row['turma']}")
                    st.write(f"**Monitor:** {row['monitor']}")
                    st.write(f"**Alunos:** {row['alunos']}")
                    st.write(f"**Relatório:** {row['relatorio']}")
                    st.markdown("</div>", unsafe_allow_html=True)