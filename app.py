import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Validador de Vendas", layout="wide")

st.title("🔍 Sistema de Conferência de Vendas - Grupo Óticas Visão")

st.markdown("Faça o upload das planilhas para validar e gerar um extrato conferido.")

uploaded_rede = st.file_uploader("📄 Envie a planilha REDE 2025", type=["xlsx"], key="rede")
uploaded_pagseguro = st.file_uploader("📄 Envie a planilha PAGSEGURO 2025", type=["xlsx"], key="pagseguro")
uploaded_extrato = st.file_uploader("📄 Envie a planilha Extrato de Vendas 2025", type=["xlsx"], key="extrato")

# Mapeamento de colunas para padronização
column_mapping = {
    "codigo_nsu": ["NSU/CV", "Código NSU", "Cód. NSU"],
    "codigo_autorizacao": ["numero da autorizaçao (Auto)", "Código de Autorização", "Codigo de Autorizacao", "AUTORIZACAO"],
    "codigo_venda": ["numero do pedido", "Código da Venda", "AUTVENDA"],
    "data": ["data da venda", "Data da Transação", "EMISSÃO"],
    "valor": ["valor da venda original", "Valor Bruto", "VALOR"],
    "loja": ["LOJA", "loja", "LOCAL"]
}

required_cols = list(column_mapping.keys())

def normalize_columns(df):
    df_copy = df.copy()
    new_columns = {}
    for logical_name, aliases in column_mapping.items():
        for col in df_copy.columns:
            if col.strip().lower() in [a.lower() for a in aliases]:
                new_columns[col] = logical_name
                break
    df_copy = df_copy.rename(columns=new_columns)
    return df_copy

def conferir_linha(linha, vendas_df):
    filtro = (
        (vendas_df["codigo_nsu"] == linha["codigo_nsu"]) &
        (vendas_df["codigo_autorizacao"] == linha["codigo_autorizacao"]) &
        (vendas_df["codigo_venda"] == linha["codigo_venda"]) &
        (vendas_df["data"] == linha["data"]) &
        (vendas_df["valor"] == linha["valor"]) &
        (vendas_df["loja"] == linha["loja"])
    )
    return "Conferido" if vendas_df[filtro].shape[0] > 0 else "Erro"

def gerar_excel_colorido(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Extrato Conferido")
        workbook = writer.book
        worksheet = writer.sheets["Extrato Conferido"]

        verde = workbook.add_format({"bg_color": "#C6EFCE"})
        vermelho = workbook.add_format({"bg_color": "#FFC7CE"})

        status_col = df.columns.get_loc("Status Conferência")

        for row_num, status in enumerate(df["Status Conferência"], start=1):
            formato = verde if status == "Conferido" else vermelho
            worksheet.set_row(row_num, cell_format=formato)

    output.seek(0)
    return output

if uploaded_rede and uploaded_pagseguro and uploaded_extrato:
    with st.spinner("Processando arquivos..."):

        # Leitura e normalização
        extrato_df = normalize_columns(pd.read_excel(uploaded_extrato, dtype=str))
        rede_df = normalize_columns(pd.read_excel(uploaded_rede, dtype=str))
        
        pagseguro_xls = pd.ExcelFile(uploaded_pagseguro)
        pagseguro_df = pd.concat(
            [normalize_columns(pd.read_excel(uploaded_pagseguro, sheet_name=s, dtype=str)) for s in pagseguro_xls.sheet_names],
            ignore_index=True
        )

        # Combinar REDE + PAGSEGURO
        vendas_df = pd.concat([rede_df, pagseguro_df], ignore_index=True)

        # Verificar colunas mínimas
        if not all(col in extrato_df.columns for col in required_cols):
            st.error("❌ A planilha de Extrato não contém todas as colunas obrigatórias.")
        elif not all(col in vendas_df.columns for col in required_cols):
            st.error("❌ As planilhas de vendas não contêm todas as colunas obrigatórias.")
        else:
            # Conferir
            extrato_df["Status Conferência"] = extrato_df.apply(lambda row: conferir_linha(row, vendas_df), axis=1)
            output_file = gerar_excel_colorido(extrato_df)

            st.success("✅ Conferência finalizada com sucesso!")
            st.dataframe(extrato_df)

            st.download_button(
                label="⬇️ Baixar Extrato Conferido (Excel)",
                data=output_file,
                file_name="Extrato_Conferido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("📥 Envie todos os arquivos para iniciar a conferência.")
