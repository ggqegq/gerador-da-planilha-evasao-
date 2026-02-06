import streamlit as st
import pandas as pd
from io import BytesIO

# Função para processar relatório por curso e modalidade
def processar_por_curso_modalidade(df, curso, periodo):
    df["Curso"] = curso  # força a coluna Curso
    df_curso = df[df["Curso"] == curso]

    modalidades = df_curso["Modalidade de Ingresso"].unique()
    resultados = []

    for modalidade in modalidades:
        df_mod = df_curso[df_curso["Modalidade de Ingresso"] == modalidade]

        metrics = {
            "Período": periodo,
            "Curso": curso,
            "Modalidade": modalidade,
            "Ingressantes": len(df_mod),
            "Cancelamentos": df_mod["Situação"].str.contains("Cancelamento", case=False, na=False).sum(),
            "Matrículas Ativas": df_mod["Situação"].str.contains("Pendente|Ativo", case=False, na=False).sum(),
            "Formados": df_mod["Situação"].str.contains("Formado", case=False, na=False).sum(),
            "% Evasão": 0.0
        }

        if metrics["Ingressantes"] > 0:
            metrics["% Evasão"] = round((metrics["Cancelamentos"] / metrics["Ingressantes"]) * 100, 2)

        resultados.append(metrics)

    return resultados

# Interface Streamlit
st.title("Gerador de Planilha Acumulada de Evasão - Química")

st.write("Carregue os relatórios separados por curso:")

files_lic = st.file_uploader("Arquivos da Licenciatura Química", type=["xlsx"], accept_multiple_files=True)
files_bach = st.file_uploader("Arquivos do Bacharel Química", type=["xlsx"], accept_multiple_files=True)
files_ind = st.file_uploader("Arquivos do Bacharel Química Industrial", type=["xlsx"], accept_multiple_files=True)

if st.button("Gerar Planilha Consolidada"):
    resultados_por_curso = {
        "Licenciatura Química": [],
        "Bacharel Química": [],
        "Bacharel Q Industrial": []
    }

    # Processa cada conjunto de arquivos
    for uploaded_file in files_lic:
        df = pd.read_excel(uploaded_file)
        periodo = uploaded_file.name.split("_")[-1].replace(".xlsx", "")
        resultados_por_curso["Licenciatura Química"].extend(processar_por_curso_modalidade(df, "Licenciatura Química", periodo))

    for uploaded_file in files_bach:
        df = pd.read_excel(uploaded_file)
        periodo = uploaded_file.name.split("_")[-1].replace(".xlsx", "")
        resultados_por_curso["Bacharel Química"].extend(processar_por_curso_modalidade(df, "Bacharel Química", periodo))

    for uploaded_file in files_ind:
        df = pd.read_excel(uploaded_file)
        periodo = uploaded_file.name.split("_")[-1].replace(".xlsx", "")
        resultados_por_curso["Bacharel Q Industrial"].extend(processar_por_curso_modalidade(df, "Bacharel Q Industrial", periodo))

    # Cria Excel com abas por curso
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for curso, dados in resultados_por_curso.items():
            if dados:
                df_resultado = pd.DataFrame(dados)
                df_resultado.to_excel(writer, sheet_name=curso, index=False)

    st.success("Planilha acumulada gerada com sucesso!")

    st.download_button(
        label="Baixar Planilha Acumulada (.xlsx)",
        data=output.getvalue(),
        file_name="evasao_acumulada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
