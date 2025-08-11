import streamlit as st
import pandas as pd
from io import BytesIO

# ================== CONFIGURAÇÃO ==================
st.set_page_config(page_title="Conferência de Vendas", layout="wide")

# Logo no topo
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 20px;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/a/a7/OOjs_UI_icon_check-ltr-progressive.svg/1200px-OOjs_UI_icon_check-ltr-progressive.svg.png" width="120">
        <h1 style="color:#2E86C1; font-family: Arial, sans-serif;">Sistema de Conferência de Vendas</h1>
        <h3 style="color:#555; font-weight: normal;">Grupo Óticas Visão</h3>
    </div>
    """,
    unsafe_allow_html=True
)

# ================== MAPEAMENTO ==================
colunas_equivalentes = {
    "codigo_nsu": ["código nsu", "nsu", "código", "codigo"],
    "autorizacao": ["código de autorizacao", "autorizacao", "autorização"],
    "codigo_venda": ["código da venda", "cod venda", "codigo venda", "codigo da venda"],
    "data": ["data", "data venda", "data da venda", "emissão"],
    "valor": ["valor", "valor bruto", "valor da venda", "valor original"],
    "loja": ["loja", "local", "unidade"]
}

# ================== FUNÇÕES ==================
def normalizar_colunas(df):
    novas_colunas = {}
    for col in df.columns:
        col_formatada = col.strip().lower()
        for chave, similares in colunas_equivalentes.items():
            if col_formatada in similares:
                novas_colunas[col] = chave
                break
    return df.rename(columns=novas_colunas)

def conferir_vendas(dfs):
    dfs_normalizados = [normalizar_colunas(df) for df in dfs]
    extrato = dfs_normalizados[0]
    outros = dfs_normalizados[1:]

    for i, df in enumerate(dfs_normalizados):
        df["status"] = "Erro"

    # Conferência cruzada
    for idx, row in extrato.iterrows():
        for df in outros:
            match = df[
                (df["data"] == row.get("data")) &
                (df["valor"] == row.get("valor")) &
                (df["loja"] == row.get("loja"))
            ]
            if not match.empty:
                extrato.at[idx, "status"] = "Conferido"

    # Atualiza status nos outros arquivos
    for df in outros:
        for idx, row in df.iterrows():
            match = extrato[
                (extrato["data"] == row.get("data")) &
                (extrato["valor"] == row.get("valor")) &
                (extrato["loja"] == row.get("loja")) &
                (extrato["status"] == "Conferido")
            ]
            if not match.empty:
                df.at[idx, "status"] = "Conferido"

    return dfs_normalizados

def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Conferência")
        workbook = writer.book
        worksheet = writer.sheets["Conferência"]

        verde = workbook.add_format({"bg_color": "#C6EFCE"})
        vermelho = workbook.add_format({"bg_color": "#FFC7CE"})

        col_valor = df.columns.get_loc("valor")
        col_status = df.columns.get_loc("status")

        for idx, status in enumerate(df["status"], start=1):
            fmt = verde if status == "Conferido" else vermelho
            worksheet.write(idx, col_valor, df.at[idx-1, "valor"], fmt)
            worksheet.write(idx, col_status, df.at[idx-1, "status"], fmt)

    output.seek(0)
    return output

# ================== SIDEBAR ==================
with st.sidebar:
    st.markdown("### 📂 Upload de Arquivos")
    extrato_file = st.file_uploader("Extrato de Vendas", type=["xlsx", "csv"])
    pagseguro_file = st.file_uploader("PAGSEGURO", type=["xlsx", "csv"])
    rede_file = st.file_uploader("REDE", type=["xlsx", "csv"])

# ================== PROCESSAMENTO ==================
if extrato_file and (pagseguro_file or rede_file):
    dfs = [pd.read_excel(extrato_file)]
    if pagseguro_file:
        dfs.append(pd.read_excel(pagseguro_file))
    if rede_file:
        dfs.append(pd.read_excel(rede_file))

    dfs_resultado = conferir_vendas(dfs)

    # Resumo no painel lateral
    total_conferido = sum(dfs_resultado[0]["status"] == "Conferido")
    total_erro = sum(dfs_resultado[0]["status"] != "Conferido")

    with st.sidebar:
        st.markdown("### 📊 Resumo da Conferência")
        st.metric("Vendas Conferidas ✅", total_conferido)
        st.metric("Vendas Não Conferidas ❌", total_erro)

    # Exibição e download para cada planilha
    for i, df_res in enumerate(dfs_resultado):
        nome = ["Extrato de Vendas", "PagSeguro", "Rede"][i]
        st.subheader(f"📄 {nome}")
        st.dataframe(df_res, use_container_width=True)

        output = exportar_excel(df_res)
        st.download_button(
            label=f"💾 Baixar {nome} Conferido",
            data=output,
            file_name=f"{nome.replace(' ', '_')}_Conferido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Faça upload do Extrato e pelo menos uma das outras planilhas (PagSeguro ou Rede).")
