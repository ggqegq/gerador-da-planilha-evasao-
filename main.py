import streamlit as st
import pandas as pd
from io import BytesIO

# Função para processar relatório por curso e modalidade
def processar_por_curso_modalidade(df, curso, periodo):
    df_curso = df[df["Curso"].str.contains(curso, case=False, na=False)]

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

uploaded_files = st.file_uploader(
    "Carregue múltiplos relatórios de listagem de alunos (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    cursos = ["Licenciatura Química", "Bacharel Química", "Bacharel Q Industrial"]
    resultados_por_curso = {curso: [] for curso in cursos}

    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)

        # Extrai período do nome do arquivo (ex: Relatório_2025.1.xlsx)
        periodo = uploaded_file.name.split("_")[-1].replace(".xlsx", "")

        for curso in cursos:
            if "Curso" in df.columns and "Modalidade de Ingresso" in df.columns:
                resultados = processar_por_curso_modalidade(df, curso, periodo)
                resultados_por_curso[curso].extend(resultados)

    # Cria Excel com abas por curso
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for curso, dados in resultados_por_curso.items():
            if dados:  # só cria aba se houver dados
                df_resultado = pd.DataFrame(dados)
                df_resultado.to_excel(writer, sheet_name=curso, index=False)

    st.success("Planilha acumulada gerada com sucesso!")

    st.download_button(
        label="Baixar Planilha Acumulada (.xlsx)",
        data=output.getvalue(),
        file_name="evasao_acumulada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
