import pandas as pd

arquivo_entrada = 'estoque_desbloqueado.xlsx'
arquivo_saida_json = 'estoque_tratado.json'

def tratar_e_salvar_em_json(arquivo_entrada, arquivo_saida_json):
    try:
        df = pd.read_excel(arquivo_entrada, header=None)

        linhas_selecionadas = list(range(3, 69)) + list(range(73, 109))
        colunas_relevantes = [0, 1, 6, 8, 9, 12, 13, 14, 15, 16, 17, 18]

        df_selecionado = df.iloc[linhas_selecionadas, colunas_relevantes]

        df_selecionado.columns = [
            "Código", "Produto", "Estoque Físico", "Estoque Disponível", "Quant. Vendida",
            "Sugestão Compras", "Lj Sugestão", "Última Venda", "Lj Última Venda",
            "Dados Compra", "Dados Compra Valor", "Dados Compra Data"
        ]

        df_selecionado = df_selecionado.dropna(how='all')
    
        df_selecionado.info(), df_selecionado.head()

        df_selecionado.to_json(arquivo_saida_json, orient='records', lines=True)
        return f"Dados tratados e salvos em: {arquivo_saida_json}"

    except Exception as e:
        return f"Erro ao processar o arquivo: {e}"

resultado = tratar_e_salvar_em_json(arquivo_entrada, arquivo_saida_json)
resultado