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

def normalizar_colunas(df):
    novas_colunas = {}
    for col in df.columns:
        col_formatada = col.strip().lower()
        for chave, similares in colunas_equivalentes.items():
            if col_formatada in similares:
                novas_colunas[col] = chave
                break
    return df.rename(columns=novas_colunas)

# Conferência bidirecional
def conferir_bidirecional(extrato, outros):
    extrato = normalizar_colunas(extrato)
    extrato["status"] = "Não conferido"

    outros_atualizados = []

    for df in outros:
        df = normalizar_colunas(df)
        df["status"] = "Não conferido"

        # Conferir Extrato -> Outro
        for idx, row in extrato.iterrows():
            match = df[
                (df["data"] == row.get("data")) &
                (df["valor"] == row.get("valor")) &
                (df["loja"] == row.get("loja"))
            ]
            if not match.empty:
                extrato.at[idx, "status"] = "Conferido"

        # Conferir Outro -> Extrato
        for idx, row in df.iterrows():
            match = extrato[
                (extrato["data"] == row.get("data")) &
                (extrato["valor"] == row.get("valor")) &
                (extrato["loja"] == row.get("loja"))
            ]
            if not match.empty:
                df.at[idx, "status"] = "Conferido"

        outros_atualizados.append(df)

    return extrato, outros_atualizados

# Exportar Excel com cores apenas nas colunas valor e status
def exportar_excel(df, sheet_name="Planilha"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        verde = workbook.add_format({"bg_color": "#C6EFCE"})
        vermelho = workbook.add_format({"bg_color": "#FFC7CE"})

        col_valor = df.columns.get_loc("valor")
        col_status = df.columns.get_loc("status")

        for idx, status in enumerate(df["status"], start=1):
            fmt = verde if status == "Conferido" else vermelho
            worksheet.write(idx, col_valor, df.iloc[idx-1, col_valor], fmt)
            worksheet.write(idx, col_status, df.iloc[idx-1, col_status], fmt)

    output.seek(0)
    return output

# ---------------------- STREAMLIT APP ----------------------
st.set_page_config(page_title="Conferência de Vendas", layout="wide")
st.title("📊 Sistema de Conferência de Vendas - Grupo Óticas Visão")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/OOjs_UI_icon_check-ltr-progressive.svg/1200px-OOjs_UI_icon_check-ltr-progressive.svg.png", width=100)
    extrato_file = st.file_uploader("Extrato de Vendas", type=["xlsx", "csv"])
    pagseguro_file = st.file_uploader("PAGSEGURO", type=["xlsx", "csv"])
    rede_file = st.file_uploader("REDE", type=["xlsx", "csv"])

if extrato_file and (pagseguro_file or rede_file):
    df_extrato = pd.read_excel(extrato_file)
    dfs_comparacao = []
    nomes_comparacao = []

    if pagseguro_file:
        dfs_comparacao.append(pd.read_excel(pagseguro_file))
        nomes_comparacao.append("PagSeguro")
    if rede_file:
        dfs_comparacao.append(pd.read_excel(rede_file))
        nomes_comparacao.append("Rede")

    st.success("Arquivos carregados com sucesso!")

    df_extrato_atualizado, outros_atualizados = conferir_bidirecional(df_extrato, dfs_comparacao)

    # Resumo na barra lateral
    with st.sidebar:
        total_conf = (df_extrato_atualizado["status"] == "Conferido").sum()
        total_nao_conf = (df_extrato_atualizado["status"] != "Conferido").sum()
        st.markdown(f"**Extrato:** {total_conf} conferidos / {total_nao_conf} não conferidos")
        for nome, df in zip(nomes_comparacao, outros_atualizados):
            total_conf = (df["status"] == "Conferido").sum()
            total_nao_conf = (df["status"] != "Conferido").sum()
            st.markdown(f"**{nome}:** {total_conf} conferidos / {total_nao_conf} não conferidos")

    # Exibir e exportar cada planilha
    st.subheader("Resultado da Conferência - Extrato")
    st.dataframe(df_extrato_atualizado, use_container_width=True)
    st.download_button(
        label="⬇ Baixar Extrato Conferido",
        data=exportar_excel(df_extrato_atualizado, "Extrato"),
        file_name="Extrato_Conferido.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    for nome, df in zip(nomes_comparacao, outros_atualizados):
        st.subheader(f"Resultado da Conferência - {nome}")
        st.dataframe(df, use_container_width=True)
        st.download_button(
            label=f"⬇ Baixar {nome} Conferido",
            data=exportar_excel(df, nome),
            file_name=f"{nome}_Conferido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Faça upload do Extrato e pelo menos uma das outras planilhas (PagSeguro ou Rede).")
