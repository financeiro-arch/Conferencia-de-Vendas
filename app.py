import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conferência de Vendas", page_icon="📊")

st.title("📋 Sistema de Conferência de Vendas - Grupo Óticas Visão")

uploaded_file1 = st.file_uploader("📎 Enviar Planilha 1 (Ex: PagSeguro)", type=["csv", "xlsx"])
uploaded_file2 = st.file_uploader("📎 Enviar Planilha 2 (Ex: Extrato)", type=["csv", "xlsx"])

if uploaded_file1 and uploaded_file2:
    df1 = pd.read_excel(uploaded_file1) if uploaded_file1.name.endswith('.xlsx') else pd.read_csv(uploaded_file1)
    df2 = pd.read_excel(uploaded_file2) if uploaded_file2.name.endswith('.xlsx') else pd.read_csv(uploaded_file2)

    # Comparação básica
    conferidas = df1.merge(df2, on=["Código NSU", "Codigo de Autorizacao", "Código da Venda"], how="inner")

    st.success(f"🔍 {len(conferidas)} vendas conferidas!")
    st.dataframe(conferidas)

    # Exportar resultado
    if not conferidas.empty:
        export = st.radio("Exportar resultado como:", ["Excel", "PDF"])
        if export == "Excel":
            from io import BytesIO
            output = BytesIO()
            conferidas.to_excel(output, index=False)
            st.download_button("⬇️ Baixar Excel", output.getvalue(), file_name="conferidas.xlsx")
