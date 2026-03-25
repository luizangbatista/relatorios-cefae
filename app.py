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

ARQUIVO_DADOS = "dados_monitoria.xlsx"

MONITORES = [
    "Luiza Matemática",
    "Gabriel Português",
]

COLUNAS_ALUNOS = ["turma", "aluno"]
COLUNAS_RELATORIOS = ["data", "turma", "monitor", "alunos", "relatorio"]

PDF_TITULO = "Relatório CEFAE"
PDF_FONTE_CORPO = "Helvetica"
PDF_FONTE_CORPO_NEGRITO = "Helvetica-Bold"
PDF_TAMANHO_CORPO = 11
PDF_TAMANHO_CABECALHO = 14

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
        max-width: 950px;
    }

    .stButton > button {
        width: 100%;
        border-radius: 12px;
        min-height: 52px;
        font-size: 16px;
    }

    .caixa {
        padding: 1rem;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        margin-bottom: 1rem;
        background-color: #FFFFFF;
    }

    .titulo-home {
        text-align: center;
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }

    .subtitulo-home {
        text-align: center;
        font-size: 1.05rem;
        color: #555;
        margin-bottom: 1.5rem;
    }

    .sucesso-home {
        padding: 0.9rem 1rem;
        border-radius: 12px;
        background-color: #ecfdf3;
        border: 1px solid #bbf7d0;
        color: #166534;
        margin-bottom: 1rem;
        text-align: center;
        font-weight: 600;
    }

    .bloco-acoes {
        padding: 0.9rem;
        border: 1px solid #E5E7EB;
        border-radius: 12px;
        background: #FAFAFA;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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


def gerar_pdf_relatorios(df, filtros_texto):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    largura, altura = A4
    margem_esq = 45
    margem_dir = 45
    y = altura - 45
    largura_texto = largura - margem_esq - margem_dir

    def nova_pagina():
        nonlocal y
        c.showPage()
        y = altura - 45

    def escrever_linha(texto, fonte=PDF_FONTE_CORPO, tamanho=PDF_TAMANHO_CORPO, espaco=16):
        nonlocal y
        linhas = simpleSplit(str(texto), fonte, tamanho, largura_texto)
        c.setFont(fonte, tamanho)
        for linha in linhas:
            if y < 60:
                nova_pagina()
                c.setFont(fonte, tamanho)
            c.drawString(margem_esq, y, linha)
            y -= espaco

    def linha_separadora():
        nonlocal y
        if y < 70:
            nova_pagina()
        c.line(margem_esq, y, largura - margem_dir, y)
        y -= 14

    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")

    escrever_linha(PDF_TITULO, PDF_FONTE_CORPO_NEGRITO, PDF_TAMANHO_CABECALHO, 20)
    escrever_linha(filtros_texto, PDF_FONTE_CORPO, PDF_TAMANHO_CORPO, 16)
    escrever_linha(data_geracao, PDF_FONTE_CORPO, PDF_TAMANHO_CORPO, 18)

    if df.empty:
        linha_separadora()
        escrever_linha("Nenhum relatório encontrado.", PDF_FONTE_CORPO_NEGRITO, PDF_TAMANHO_CORPO, 16)
    else:
        for _, row in df.iterrows():
            try:
                data_formatada = pd.to_datetime(row.get("data", ""), errors="coerce").strftime("%d/%m/%Y")
            except Exception:
                data_formatada = str(row.get("data", ""))

            linha_separadora()
            escrever_linha(f"Data: {data_formatada}")
            escrever_linha(f"Turma: {row.get('turma', '')}")
            escrever_linha(f"Monitor: {row.get('monitor', '')}")
            escrever_linha(f"Alunos: {row.get('alunos', '')}")
            escrever_linha("Relatório:", PDF_FONTE_CORPO_NEGRITO, PDF_TAMANHO_CORPO, 16)
            escrever_linha(f"{row.get('relatorio', '')}", PDF_FONTE_CORPO, PDF_TAMANHO_CORPO, 18)

    c.save()
    buffer.seek(0)
    return buffer


def gerar_docx_relatorios(df, filtros_texto):
    doc = Document()

    estilo_normal = doc.styles["Normal"]
    estilo_normal.font.name = "Calibri"
    estilo_normal.font.size = Pt(11)

    titulo = doc.add_paragraph()
    run_titulo = titulo.add_run("Relatório CEFAE")
    run_titulo.bold = True
    run_titulo.font.name = "Calibri"
    run_titulo.font.size = Pt(14)

    p_filtros = doc.add_paragraph()
    r_filtros = p_filtros.add_run(filtros_texto)
    r_filtros.font.name = "Calibri"
    r_filtros.font.size = Pt(11)

    p_data = doc.add_paragraph()
    r_data = p_data.add_run(datetime.now().strftime("%d/%m/%Y %H:%M"))
    r_data.font.name = "Calibri"
    r_data.font.size = Pt(11)

    if df.empty:
        p = doc.add_paragraph()
        r = p.add_run("Nenhum relatório encontrado.")
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(11)
    else:
        for _, row in df.iterrows():
            try:
                data_formatada = pd.to_datetime(row.get("data", ""), errors="coerce").strftime("%d/%m/%Y")
            except Exception:
                data_formatada = str(row.get("data", ""))

            doc.add_paragraph("_" * 70)

            for rotulo, valor in [
                ("Data", data_formatada),
                ("Turma", row.get("turma", "")),
                ("Monitor", row.get("monitor", "")),
                ("Alunos", row.get("alunos", "")),
            ]:
                p = doc.add_paragraph()
                r1 = p.add_run(f"{rotulo}: ")
                r1.bold = True
                r1.font.name = "Calibri"
                r1.font.size = Pt(11)

                r2 = p.add_run(str(valor))
                r2.font.name = "Calibri"
                r2.font.size = Pt(11)

            p_rel = doc.add_paragraph()
            r_rel_t = p_rel.add_run("Relatório: ")
            r_rel_t.bold = True
            r_rel_t.font.name = "Calibri"
            r_rel_t.font.size = Pt(11)

            r_rel_v = p_rel.add_run(str(row.get("relatorio", "")))
            r_rel_v.font.name = "Calibri"
            r_rel_v.font.size = Pt(11)

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


def ir_para(nome_pagina):
    st.session_state.pagina = nome_pagina
    st.rerun()


def botao_voltar():
    if st.button("⬅️ Voltar para a página inicial"):
        st.session_state.modo_exclusao = False
        ir_para("home")


def tela_home():
    st.markdown('<div class="titulo-home">📚 Sistema de Monitoria</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitulo-home">Selecione uma das opções abaixo</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.mensagem_sucesso:
        st.markdown(
            f'<div class="sucesso-home">{st.session_state.mensagem_sucesso}</div>',
            unsafe_allow_html=True,
        )
        st.session_state.mensagem_sucesso = ""

    st.write("")
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


def texto_filtros_pdf_docx(turma_filtro, aluno_filtro, monitor_filtro, data_ini, data_fim):
    return " | ".join([
        f"Turma: {turma_filtro}",
        f"Aluno: {aluno_filtro}",
        f"Monitor: {monitor_filtro}",
        f"Data inicial: {data_ini.strftime('%d/%m/%Y') if data_ini else '-'}",
        f"Data final: {data_fim.strftime('%d/%m/%Y') if data_fim else '-'}",
    ])


inicializar_arquivo()

df_alunos = carregar_alunos()
df_relatorios = carregar_relatorios()
pagina = st.session_state.pagina

if pagina == "home":
    tela_home()

elif pagina == "cadastrar_turma":
    botao_voltar()
    st.title("Cadastrar turma")
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

        st.write("**Alunos cadastrados nessa turma:**")
        for aluno in alunos_turma:
            st.write(f"- {aluno}")

elif pagina == "cadastrar_relatorio":
    botao_voltar()
    st.title("Enviar novo relatório")
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

elif pagina == "consultar":
    botao_voltar()
    st.title("Consultar relatórios enviados")
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

        filtros_texto = texto_filtros_pdf_docx(
            turma_filtro, aluno_filtro, monitor_filtro, data_ini, data_fim
        )

        st.markdown(f"**Total encontrado:** {len(df_filtrado)} relatório(s)")

        if df_filtrado.empty:
            st.warning("Nenhum relatório corresponde aos filtros selecionados.")
            st.session_state.modo_exclusao = False
        else:
            st.markdown('<div class="bloco-acoes">', unsafe_allow_html=True)

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

            st.markdown("</div>", unsafe_allow_html=True)

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

            if st.session_state.modo_exclusao:
                st.markdown("### Selecione os relatórios para excluir")

            indices_para_excluir = []

            for idx, row in df_filtrado.iterrows():
                try:
                    data_formatada = pd.to_datetime(row["data"]).strftime("%d/%m/%Y")
                except Exception:
                    data_formatada = str(row["data"])

                st.markdown('<div class="caixa">', unsafe_allow_html=True)

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