import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Funções auxiliares
def calcular_BAL_individual(dap):
    return (np.pi * (dap ** 2)) / 40000

def calcular_BAL(df, DAP_i):
    return round(sum(calcular_BAL_individual(dap) for dap in df['DAP'] if dap > DAP_i), 4)

def calcular_BAI(dap, daps_vizinhos):
    Gi = calcular_BAL_individual(dap)
    if len(daps_vizinhos) == 0:
        return np.nan
    Q = np.sqrt(sum(d**2 for d in daps_vizinhos) / len(daps_vizinhos))
    Gq = (np.pi * Q**2) / 40000
    return round(Gi / Gq, 4) if Gq != 0 else np.nan

def calcular_IC1(dap, altura):
    return round(dap / altura, 4) if altura != 0 else np.nan

def calcular_IC2(altura, altura_media):
    return round(altura / altura_media, 4) if altura_media != 0 else np.nan

def calcular_IC3(ic1, ic2):
    return round(ic1 * ic2, 4) if not np.isnan(ic1) and not np.isnan(ic2) else np.nan

def calcular_IC4(dap, altura, media_dap_altura):
    if altura == 0 or media_dap_altura == 0:
        return np.nan
    return round((dap / altura) / media_dap_altura, 4)

# Criar modelo Excel para download
def criar_modelo_excel():
    df_modelo = pd.DataFrame(columns=['Codigo_Parcela', 'Ano_Medicao', 'DAP', 'Altura', 'Especie'])
    output = BytesIO()
    df_modelo.to_excel(output, index=False)
    return output.getvalue()

# Início do Streamlit
st.title("Simulador de Índices de Competição Florestal")

modelo_excel = criar_modelo_excel()
st.download_button(
    label="📥 Baixe o modelo da planilha Excel",
    data=modelo_excel,
    file_name="modelo_planilha.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("""
### Instruções:
- A planilha deve conter colunas com os nomes exatos (maiúsculo):
  - Codigo_Parcela
  - Ano_Medicao
  - DAP
  - Altura
  - Especie
- Árvores com DAP ou Altura igual a 0 serão desconsideradas dos cálculos.
""")

uploaded_file = st.file_uploader("📂 Envie sua planilha Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.write("Colunas encontradas no arquivo:", list(df.columns))

    df.columns = df.columns.str.strip()

    colunas_esperadas = {'Codigo_Parcela', 'Ano_Medicao', 'DAP', 'Altura', 'Especie'}
    colunas_encontradas = set(df.columns)

    if not colunas_esperadas.issubset(colunas_encontradas):
        st.error("⚠️ Colunas obrigatórias faltando na planilha.")
        st.stop()

    st.success("✅ Planilha carregada corretamente.")

    df = df[(df['DAP'] > 0) & (df['Altura'] > 0)].copy()

    parcelas = df['Codigo_Parcela'].unique()
    parcela_escolhida = st.selectbox("🔹 Selecione a parcela:", parcelas)

    anos = df[df['Codigo_Parcela'] == parcela_escolhida]['Ano_Medicao'].unique()
    ano_escolhido = st.selectbox("📅 Selecione o ano de medição:", anos)

    opcao = st.selectbox("📌 Escolha o índice de competição:", ["IC1", "IC2", "IC3", "IC4", "BAL", "BAI"])

    if st.button("▶️ Calcular"):
        parcela = df[(df['Codigo_Parcela'] == parcela_escolhida) & (df['Ano_Medicao'] == ano_escolhido)].copy()
        parcela = parcela.reset_index(drop=True)
        parcela["Número da Árvore"] = parcela.index + 1

        media_dap = parcela['DAP'].mean()
        media_altura = parcela['Altura'].mean()
        media_dap_altura = (parcela['DAP'] / parcela['Altura']).mean()

        resultados = []
        for i, row in parcela.iterrows():
            dap = row['DAP']
            altura = row['Altura']
            especie = row['Especie']
            num_arvore = row['Número da Árvore']

            # Remove árvore atual das médias de vizinhança
            vizinhos = parcela[parcela['Número da Árvore'] != num_arvore]
            daps_vizinhos = vizinhos['DAP'].tolist()
            altura_vizinhos = vizinhos['Altura'].tolist()

            ic1 = calcular_IC1(dap, altura)
            ic2 = calcular_IC2(altura, np.mean(altura_vizinhos))
            ic3 = calcular_IC3(ic1, ic2)
            ic4 = calcular_IC4(dap, altura, media_dap_altura)
            bal = calcular_BAL(vizinhos, dap)
            bai = calcular_BAI(dap, daps_vizinhos)

            resultado = {
                "Número da Árvore": num_arvore,
                "Codigo_Parcela": parcela_escolhida,
                "Ano_Medicao": ano_escolhido,
                "Especie": especie,
                "DAP": dap,
                "Altura": altura
            }

            if opcao == "IC1":
                resultado["IC1"] = ic1
            elif opcao == "IC2":
                resultado["IC2"] = ic2
            elif opcao == "IC3":
                resultado["IC3"] = ic3
            elif opcao == "IC4":
                resultado["IC4"] = ic4
            elif opcao == "BAL":
                resultado["BAL"] = bal
            elif opcao == "BAI":
                resultado["BAI"] = bai

            resultados.append(resultado)

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
