# Curva-A-B-C---Estoque
O relatório precisa ter o nome do mês e estar na mesma pasta do código, desta forma selecionando os filtros que deseja, o código gerar um arquivo com os itens de acordo com os filtros utilizados.

import pandas as pd

# Variáveis para Filtro
curva_desejada = "A"             # Ex: "A", "B", "C"
uso_desejado = "Imobilizado"     # Ex: "Imobilizado", "Consumo", "Outros insumos"
mes = "Junho"

# Caminhos dos arquivos
caminho = f"/content/drive/MyDrive/Curva ABC/Curva ABC {mes}.xls"
caminho_usado_como = "/content/drive/MyDrive/Curva ABC/Classificação, usado como.xlsx"
caminho_curva_final = f"/content/drive/MyDrive/Curva ABC/Curva {curva_desejada} {uso_desejado} {mes}.xlsx"


# Função para carregar e classificar os dados
def carregar_e_classificar(caminho, caminho_usado_como):
    df_raw = pd.read_excel(caminho, header=None)
    linha_cabecalho = df_raw[df_raw.eq('Codigo Empresa').any(axis=1)].index[0]
    df = df_raw.iloc[linha_cabecalho + 1:].copy()
    df.columns = df_raw.iloc[linha_cabecalho]
    df = df.dropna(axis=1, how='all').dropna(how='any').reset_index(drop=True)

    # Conversões
    df["Qtd"] = pd.to_numeric(df["Qtd"], errors='coerce')
    df["Valor da Venda"] = pd.to_numeric(df["Valor da Venda"], errors='coerce')
    df["Valor Total"] = df["Qtd"] * df["Valor da Venda"]

    # Agrupamento
    df_produtos = df.groupby("Cod.Prod.").agg({
        "Descricao do Produto": "first",
        "Qtd": "sum",
        "Valor da Venda": "mean",
        "Valor Total": "sum"
    }).reset_index()

    # Junta com classificação de uso
    df_usado_como = pd.read_excel(caminho_usado_como)
    df_completo = pd.merge(df_produtos, df_usado_como, on="Cod.Prod.", how="left")

    # Agrupa por uso e produto
    df_agrupado = df_completo.groupby(["Usado como", "Cod.Prod."]).agg({
        "Descricao do Produto": "first",
        "Qtd": "sum",
        "Valor da Venda": "mean",
        "Valor Total": "sum"
    }).reset_index()

    # Curva ABC
    def aplicar_curva_abc(grupo):
        grupo = grupo.sort_values("Qtd", ascending=False).reset_index(drop=True)
        total = grupo["Valor Total"].sum()
        grupo["% Acumulado"] = 100 * grupo["Valor Total"].cumsum() / total
        grupo["Classe ABC"] = grupo["% Acumulado"].apply(
            lambda p: 'A' if p <= 80 else 'B' if p <= 95 else 'C')
        return grupo

    df_classificado = df_agrupado.groupby("Usado como", group_keys=False).apply(aplicar_curva_abc)
    return df_classificado

# Carrega e classifica
df_final = carregar_e_classificar(caminho, caminho_usado_como)

# Aplica filtros combinados
df_filtrado = df_final[
    (df_final["Classe ABC"] == curva_desejada.upper()) &
    (df_final["Usado como"].str.lower() == uso_desejado.lower())
].reset_index(drop=True)

# Exibe resultado
df_filtrado

#Salvando o dataframe
df_filtrado.to_excel(caminho_curva_final, index=False)
