import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Funções auxiliares
def calcular_BAL_individual(dap):
    return (3.1416 * (dap ** 2)) / 40000

def calcular_BAI(dap):
    return round((3.1416 * (dap / 2) ** 2) / 10000, 4)

def calcular_BAL(df, dap_i):
    soma_BAL_maiores = sum(
        calcular_BAL_individual(dap) for dap in df['DAP'] if dap > dap_i
    )
    return round(soma_BAL_maiores, 4)

def calcular_IC1(dap, d_medio):
    return round((dap ** 2) / (d_medio ** 2), 4) if d_medio != 0 else np.nan

def calcular_IC2(altura, altura_media):
    return round(altura / altura_media, 4) if altura_media != 0 else np.nan

def calcular_IC3(dap, d_medio, altura, altura_media):
    ic1 = calcular_IC1(dap, d_medio)
    ic2 = calcular_IC2(altura, altura_media)
    return round(ic1 * ic2, 4)

def area_seccional(dap):
    return (3.1416 * dap ** 2) / 40000

def area_seccional_quadratica(d_medio):
    return (3.1416 * d_medio ** 2) / 40000

def calcular_IC4(dap, d_medio):
    AS_i = area_seccional(dap)
    ASq = area_seccional_quadratica(d_medio)
    return round((AS_i ** 2) / (ASq ** 2), 4) if ASq != 0 else np.nan

# Criar modelo Excel
def criar_modelo_excel():
    df_modelo = pd.DataFrame(columns=['Codigo_Parcela', 'Ano_Medicao', 'DAP', 'Altura', 'Espécie'])
    output = BytesIO()
    df_modelo.to_excel(output, index=False)
    return output.getvalue()

# Streamlit App
st.title("Simulador de Índices de Competição Florestal")

modelo_excel = criar_modelo_excel()
st.download_button(
    label="Baixe o modelo da planilha Excel",
    data=modelo_excel,
    file_name="modelo_planilha.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("""
### Instruções para a planilha Excel

- O arquivo deve conter as seguintes colunas:
    - Codigo_Parcela
    - Ano_Medicao
    - DAP
    - Altura
    - Espécie
""")

uploaded_file = st.file_uploader("Envie sua planilha Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.write("Colunas encontradas:", list(df.columns))

    colunas_esperadas = {'Codigo_Parcela', 'Ano_Medicao', 'DAP', 'Altura', 'Espécie'}
    colunas_encontradas = set(df.columns)

    faltando = colunas_esperadas - colunas_encontradas
    if faltando:
        st.error(f"Faltam colunas obrigatórias: {', '.join(faltando)}")
        st.stop()

    st.success("Todas as colunas obrigatórias foram encontradas.")

    parcelas = df['Codigo_Parcela'].unique()
    parcela_escolhida = st.selectbox("Selecione a parcela:", parcelas)

    opcao = st.selectbox("Escolha o índice de competição:", [
        "IC1", "IC2", "IC3", "IC4", "BAL", "BAI"
    ])

    if st.button("▶️ Calcular"):
        parcela = df[df["Codigo_Parcela"] == parcela_escolhida].copy()
        parcela = parcela.reset_index(drop=True)
        parcela["Número_Árvore"] = range(1, len(parcela) + 1)

        resultados = []
        for i, row in parcela.iterrows():
            dap = row["DAP"]
            altura = row["Altura"]
            especie = row.get("Espécie", "Desconhecida")
            num_arvore = row["Número_Árvore"]

            if pd.isnull(dap) or pd.isnull(altura):
                continue

            # Remove árvore atual para média
            sem_atual = parcela[parcela["Número_Árvore"] != num_arvore]
            d_medio = sem_atual["DAP"].mean()
            altura_media = sem_atual["Altura"].mean()

            resultado = {
                "Número da Árvore": num_arvore,
                "Código Parcela": parcela_escolhida,
                "Espécie": especie,
                "DAP": dap,
                "Altura": altura
            }

            if opcao == "IC1":
                resultado["IC1"] = calcular_IC1(dap, d_medio)
            elif opcao == "IC2":
                resultado["IC2"] = calcular_IC2(altura, altura_media)
            elif opcao == "IC3":
                resultado["IC3"] = calcular_IC3(dap, d_medio, altura, altura_media)
            elif opcao == "IC4":
                resultado["IC4"] = calcular_IC4(dap, d_medio)
            elif opcao == "BAL":
                resultado["BAL"] = calcular_BAL(parcela, dap)
            elif opcao == "BAI":
                resultado["BAI"] = calcular_BAI(dap)

            resultados.append(resultado)

        if resultados:
            df_resultados = pd.DataFrame(resultados)
            st.subheader("📊 Resultados")
            st.dataframe(df_resultados)

            output = BytesIO()
            df_resultados.to_excel(output, index=False, engine='openpyxl')
            st.download_button(
                label="⬇️ Baixar resultados em Excel",
                data=output.getvalue(),
                file_name="resultados_simulador.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Nenhum resultado calculado. Verifique os dados.")
