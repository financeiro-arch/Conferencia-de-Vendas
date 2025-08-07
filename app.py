import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Conferência de Vendas - Grupo Óticas Visão", page_icon="📊")

st.title("📋 Conferência de Vendas")
st.subheader("Grupo Óticas Visão")

st.markdown("---")
st.markdown("Envie as duas planilhas para validar as vendas com base nos campos:")
st.markdown("- Código NSU\n- Código de Autorização\n- Código da Venda")

uploaded_file1 = st.file_uploader("📎 Enviar Planilha 1 (Ex: PagSeguro)", type=["csv", "xlsx"], key="file1")
uploaded_file2 = st.file_uploader("📎 Enviar Planilha 2 (Ex: Extrato)", type=["csv", "xlsx"], key="file2")

if uploaded_file1 and uploaded_file2:
    def read_file(uploaded_file):
        if uploaded_file.name.endswith('.xlsx'):
            return pd.read_excel(uploaded_file, dtype=str)
        else:
            return pd.read_csv(uploaded_file, dtype=str)

    try:
        df1 = read_file(uploaded_file1)
        df2 = read_file(uploaded_file2)

        st.markdown("### 📄 Pré-visualização das planilhas")
        with st.expander("📘 Planilha 1"):
            st.dataframe(df1.head())
        with st.expander("📗 Planilha 2"):
            st.dataframe(df2.head())

        # Normaliza nomes de colunas
        df1.columns = df1.columns.str.strip()
        df2.columns = df2.columns.str.strip()

        chaves = ["Código NSU", "Codigo de Autorizacao", "Código da Venda"]

        if all(col in df1.columns for col in chaves) and all(col in df2.columns for col in chaves):
            conferidas = df1.merge(df2, on=chaves, how="inner")

            st.success(f"🔍 Foram encontradas {len(conferidas)} vendas conferidas!")
            st.dataframe(conferidas)

            # Exportação
            export = st.radio("Exportar resultado como:", ["Excel", "CSV"], horizontal=True)

            if export == "Excel":
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    conferidas.to_excel(writer, index=False)
                st.download_button("⬇️ Baixar Excel", output.getvalue(), file_name="conferidas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            elif export == "CSV":
                csv = conferidas.to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ Baixar CSV", csv, file_name="conferidas.csv", mime="text/csv")

        else:
            st.error("❌ As colunas obrigatórias não foram encontradas em ambas as planilhas. Verifique os nomes:")
            st.code("\n".join(chaves))

    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")

else:
    st.info("🔄 Aguarde o envio das duas planilhas para iniciar a conferência.")
