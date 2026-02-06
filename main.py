import streamlit as st
import pandas as pd
from io import BytesIO

# Função para detectar curso pela última linha
def detectar_curso(df):
    ultima_linha = " ".join(df.iloc[-1].astype(str).tolist()).lower()
    if "químico industrial" in ultima_linha:
        return "Bacharel Q Industrial"
    elif "licenciado" in ultima_linha:
        return "Licenciatura Química"
    elif "bacharel" in ultima_linha:
        return "Bacharel Química"
    else:
        return "Curso Desconhecido"

# Função para converter código da matrícula em período letivo
def detectar_periodo(matricula):
    try:
        codigo = str(matricula).split(".")[0][:3]  # pega os 3 primeiros dígitos
        ano = 2000 + int(codigo[1:])               # últimos 2 dígitos = ano
        periodo = codigo[0]                        # primeiro dígito = período (1 ou 2)
        return f"{ano}.{periodo}"
    except:
        return "Desconhecido"

# Função principal de processamento
def processar_relatorio(df):
    # Normaliza colunas
    df.columns = df.columns.str.strip().str.lower()

    # Detecta curso pela última linha
    curso = detectar_curso(df)

    # Remove a última linha (resumo)
    df = df.iloc[:-1, :]

    # Cria coluna curso
    df["curso"] = curso

    # Cria coluna período a partir da matrícula
    df["período"] = df["matrícula"].apply(detectar_periodo)

    # Cria coluna modalidade simplificada
    df["modalidade"] = df["modalidade de ingresso"].astype(str).str[0].map(
        lambda x: "AC" if x.upper().startswith("A") else "AA"
    )

    resultados = []
    for periodo in df["período"].unique():
        for modalidade in df["modalidade"].unique():
            df_mod = df[(df["período"] == periodo) & (df["modalidade"] == modalidade)]

            metrics = {
                "Período": periodo,
                "Curso": curso,
                "Modalidade": modalidade,
                "Ingressantes": len(df_mod),
                "Cancelamentos": df_mod["situação"].str.contains("cancelamento", case=False, na=False).sum(),
                "Matrículas Ativas": df_mod["situação"].str.contains("pendente|ativo", case=False, na=False).sum(),
                "Formados": df_mod["situação"].str.contains("formado", case=False, na=False).sum(),
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
    resultados_por_curso = {}

    for uploaded_file in uploaded_files:
        # Força cabeçalho na linha correta (ajuste se necessário)
        df = pd.read_excel(uploaded_file, header=5)
        df.columns = df.columns.str.strip().str.lower()
        st.write("Colunas detectadas:", df.columns.tolist())  # debug para confirmar

        resultados = processar_relatorio(df)

        for r in resultados:
            curso = r["Curso"]
            if curso not in resultados_por_curso:
                resultados_por_curso[curso] = []
            resultados_por_curso[curso].append(r)

    # Cria Excel com abas por curso
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for curso, dados in resultados_por_curso.items():
            if dados:
                df_resultado = pd.DataFrame(dados)
                df_resultado.to_excel(writer, sheet_name=curso, index=False)

        # Aba resumo geral
        todos = [r for lista in resultados_por_curso.values() for r in lista]
        df_geral = pd.DataFrame(todos)
        df_geral.to_excel(writer, sheet_name="Resumo Geral", index=False)

    st.success("Planilha acumulada gerada com sucesso!")

    st.download_button(
        label="Baixar Planilha Acumulada (.xlsx)",
        data=output.getvalue(),
        file_name="evasao_acumulada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
