import pandas as pd
from difflib import get_close_matches

# Mapeamento de nomes semelhantes
mapa_colunas = {
    'codigo da venda': ['código da venda', 'cod venda', 'codigo venda'],
    'nsu': ['nsu', 'código nsu', 'código', 'cod nsu'],
    'autorizacao': ['codigo de autorizacao', 'autorização', 'autorizacao'],
    'data': ['data', 'data da venda', 'emissão', 'data da transação'],
    'valor': ['valor', 'valor bruto', 'valor venda', 'valor original', 'valor da venda'],
    'loja': ['loja', 'local', 'nome loja']
}

# Função para normalizar os nomes
def normalizar_colunas(df):
    novas_colunas = {}
    for col in df.columns:
        col_norm = col.strip().lower()
        for chave, similares in mapa_colunas.items():
            if col_norm in similares or get_close_matches(col_norm, similares):
                novas_colunas[col] = chave
                break
    df = df.rename(columns=novas_colunas)
    return df

# Função principal de conferência
def conferir_vendas(arquivo_extrato, arquivo_pagseguro, arquivo_rede):
    extrato = pd.read_excel(arquivo_extrato)
    pagseguro = pd.read_excel(arquivo_pagseguro)
    rede = pd.read_excel(arquivo_rede)

    extrato = normalizar_colunas(extrato)
    pagseguro = normalizar_colunas(pagseguro)
    rede = normalizar_colunas(rede)

    extrato['status'] = 'Erro'

    # Conferência por data, valor e loja
    for idx, linha in extrato.iterrows():
        data = linha.get('data')
        valor = linha.get('valor')
        loja = linha.get('loja')

        match_pag = pagseguro[
            (pagseguro['data'] == data) &
            (pagseguro['valor'] == valor) &
            (pagseguro['loja'] == loja)
        ]
        match_rede = rede[
            (rede['data'] == data) &
            (rede['valor'] == valor) &
            (rede['loja'] == loja)
        ]

        if not match_pag.empty or not match_rede.empty:
            extrato.at[idx, 'status'] = 'Conferido'

    # Salva a planilha com sombreamento colorido
    writer = pd.ExcelWriter('Extrato_Validado.xlsx', engine='xlsxwriter')
    extrato.to_excel(writer, index=False, sheet_name='Conferência')
    workbook = writer.book
    worksheet = writer.sheets['Conferência']

    format_verde = workbook.add_format({'bg_color': '#C6EFCE'})
    format_vermelho = workbook.add_format({'bg_color': '#FFC7CE'})

    for row in range(1, len(extrato) + 1):
        status = extrato.loc[row - 1, 'status']
        fmt = format_verde if status == 'Conferido' else format_vermelho
        worksheet.set_row(row, None, fmt)

    writer.close()
    print("✅ Planilha 'Extrato_Validado.xlsx' criada com sucesso!")

# Use os nomes dos arquivos reais aqui
conferir_vendas('ExtratoVendas 2025.xlsx', 'PAGSEGURO 2025.xlsx', 'REDE 2025.xlsx')
