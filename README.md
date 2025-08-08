# Sistema de Conferência de Vendas - Grupo Óticas Visão

Este é um sistema desenvolvido em Python utilizando Streamlit, que realiza a conferência automática entre planilhas de vendas de diferentes plataformas (como PagSeguro e Rede) com o extrato de vendas oficial.

## Funcionalidades

- Upload de múltiplas planilhas (Excel ou CSV)
- Mapeamento inteligente de colunas semelhantes
- Validação automática de vendas com base em:
  - Código NSU
  - Código de Autorização
  - Código da Venda
  - Data (Data, Emissão ou Data da Venda)
  - Valor (Valor, Valor Bruto ou Valor da Venda Original)
  - Loja (Loja ou Local)
- Geração de planilha final com status: **"Conferido"** ou **"Erro"**
- Exportação em Excel com coloração verde/vermelho

## Como executar localmente

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/conferencia-vendas.git
cd conferencia-vendas
```

2. Instale os pacotes:
```bash
pip install -r requirements.txt
```

3. Execute a aplicação:
```bash
streamlit run app.py
```

## Exportação

Você pode baixar o resultado final da conferência diretamente pela interface do aplicativo.

---

Desenvolvido por **Grupo Óticas Visão**