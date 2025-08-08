import streamlit as st
import pandas as pd
from io import BytesIO

# Mapeamento de colunas equivalentes
colunas_equivalentes = {
    "codigo_nsu": ["código nsu", "nsu", "código", "codigo"],
    "autorizacao": ["código de autorizacao", "autorizacao", "autorização"],
    "codigo_venda": ["código da venda", "cod venda", "codigo venda", "codigo da venda"],
    "data": ["data", "data venda", "data da venda", "emissão"],
    "valor": ["valor", "valor bruto", "valor da venda", "valor original"],
    "loja": ["loja", "local", "unidade"]
}

# Função para renomear colunas
def normalizar_colunas(df):
    novas_colunas = {}
    for col in df.columns:
        col_formatada = col.strip().lower()
        for chave, similares in colunas_equivalentes.items():
            if col_formatada in similares:
                novas_colunas[col] = chave
                break
    return df.rename(columns=novas_colunas)

# Conferência das vendas
def conferir_vendas(extrato, outros):
    extrato = normalizar_colunas(extrato)
    extrato["status"] = "Erro"
    for idx, row in extrato.iterrows():
        for df in outros:
            df = normalizar_colunas(df)
            match = df[
                (df["data"] == row.get("data")) &
                (df["valor"] == row.get("valor")) &
                (df["loja"] == row.get("loja"))
            ]
            if not match.empty:
                extrato.at[idx, "status"] = "Conferido"
                break
    return extrato

# Exportar com cor no Excel
def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Conferência")
        workbook = writer.book
        worksheet = writer.sheets["Conferência"]

        verde = workbook.add_format({"bg_color": "#C6EFCE"})
        vermelho = workbook.add_format({"bg_color": "#FFC7CE"})

        for idx, status in enumerate(df["status"], start=1):
            fmt = verde if status == "Conferido" else vermelho
            worksheet.set_row(idx, None, fmt)

    output.seek(0)
    return output

# ---------------------- STREAMLIT APP ----------------------

st.set_page_config(page_title="Conferência de Vendas", layout="wide")
st.title("📊 Sistema de Conferência de Vendas - Grupo Óticas Visão")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/OOjs_UI_icon_check-ltr-progressive.svg/1200px-OOjs_UI_icon_check-ltr-progressive.svg.png", width=100)
    st.markdown("Faça o upload dos arquivos a seguir:")

    extrato_file = st.file_uploader("Extrato de Vendas", type=["xlsx", "csv"])
    pagseguro_file = st.file_uploader("PAGSEGURO", type=["xlsx", "csv"])
    rede_file = st.file_uploader("REDE", type=["xlsx", "csv"])

if extrato_file and (pagseguro_file or rede_file):
    df_extrato = pd.read_excel(extrato_file)
    dfs_comparacao = []
    if pagseguro_file:
        dfs_comparacao.append(pd.read_excel(pagseguro_file))
    if rede_file:
        dfs_comparacao.append(pd.read_excel(rede_file))

    st.success("Arquivos carregados com sucesso!")
    df_resultado = conferir_vendas(df_extrato, dfs_comparacao)

    st.subheader("Resultado da Conferência")
    st.dataframe(df_resultado, use_container_width=True)

    output = exportar_excel(df_resultado)

    st.download_button(
        label="📅 Baixar Resultado em Excel",
        data=output,
        file_name="Extrato_Conferido.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Faça upload do Extrato e pelo menos uma das outras planilhas (PagSeguro ou Rede).")
