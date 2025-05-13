import pandas as pd
import os

def carregar_arquivo(file_path):
    """
    Carrega um arquivo com base na extensão para o formato adequado.
    Suporta .xls, .xlsx, .csv e outros formatos.
    """
    # Extrair a extensão do arquivo
    file_extension = os.path.splitext(file_path)[1].lower()

    # Verificar e carregar o arquivo conforme sua extensão
    if file_extension in ['.xls', '.xlsx']:
        # Carregar arquivos Excel
        df = pd.read_excel(file_path)
    elif file_extension == '.csv':
        # Carregar arquivos CSV
        df = pd.read_csv(file_path)
    else:
        raise ValueError(f"Formato de arquivo não suportado: {file_extension}")
    
    return df

# Exemplo de uso
arquivo = 'retail_sales_data.xlsx'  # Ou .csv ou qualquer outro formato suportado


df = carregar_arquivo(arquivo)
print(df.head())


import pandas as pd
import numpy as np
import os

# Função para carregar arquivos de diferentes formatos (Excel, CSV, JSON, etc.)
def carregar_arquivo(file_path):
    """
    Carrega um arquivo com base na extensão para o formato adequado.
    Suporta .xls, .xlsx, .csv, .json, entre outros.
    """
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    elif file_extension == '.csv':
        df = pd.read_csv(file_path)
    elif file_extension == '.json':
        df = pd.read_json(file_path)
    elif file_extension == '.html':
        df = pd.read_html(file_path)[0]  # pd.read_html retorna uma lista, então pegamos o primeiro item
    else:
        raise ValueError(f"Formato de arquivo não suportado: {file_extension}")
    
    return df

# Função para verificar e limpar os dados
def limpar_planilha(df):
    """
    Aplica as técnicas de limpeza de dados e retorna a planilha limpa.
    """
    # Remover duplicatas
    df.drop_duplicates(inplace=True)

    # Remover espaços em branco nas colunas
    if 'Produto' in df.columns:
        df['Produto'] = df['Produto'].str.strip()
    if 'Categoria' in df.columns:
        df['Categoria'] = df['Categoria'].str.strip()
    if 'Empresa' in df.columns:
        df['Empresa'] = df['Empresa'].str.strip()

    # Garantir que 'Valor_Total' seja numérico
    df['Valor_Total'] = pd.to_numeric(df['Valor_Total'], errors='coerce')

    # Garantir que a coluna 'Data_Venda' seja de data
    df['Data_Venda'] = pd.to_datetime(df['Data_Venda'], errors='coerce')

    # Preencher valores ausentes (NaN) nas colunas numéricas com a média
    df['Valor_Total'].fillna(df['Valor_Total'].mean(), inplace=True)

    # Verificar se não há mais valores ausentes nas colunas de interesse
    print("\nValores ausentes por coluna:")
    print(df.isnull().sum())

    return df

# Função para análise das categorias e produtos
def analise_categorias_produtos(df):
    """
    Análise das categorias e produtos mais vendidos.
    """
    # Verificando se as colunas 'Categoria' e 'Produto' existem no DataFrame
    if 'Categoria' not in df.columns or 'Produto' not in df.columns:
        raise ValueError("As colunas 'Categoria' e 'Produto' não foram encontradas no DataFrame.")

    # Analisando as Categorias
    categorias_soma = df.groupby('Categoria')['Valor_Total'].sum().sort_values(ascending=False)
    print("\nCategorias de Produtos Mais Vendidos:")
    print(categorias_soma)

    # Analisando os Produtos
    produtos_soma = df.groupby('Produto')['Valor_Total'].sum().sort_values(ascending=False)
    print("\nProdutos Mais Vendidos:")
    print(produtos_soma)

    return categorias_soma, produtos_soma

# Função para carregar, limpar e salvar o arquivo
def carregar_e_limpar_arquivo(caminho_arquivo, caminho_saida):
    """
    Carrega, limpa os dados e gera a análise de categorias e produtos.
    """
    # Carregar o arquivo com base no formato
    df = carregar_arquivo(caminho_arquivo)

    # Limpar os dados
    df_limpo = limpar_planilha(df)

    # Salvar o arquivo limpo em um novo caminho
    df_limpo.to_excel(caminho_saida, index=False)
    print(f"\nPlanilha limpa salva em: {caminho_saida}")
    
    return df_limpo

# Função para carregar a planilha limpa e realizar a análise
def realizar_analise(caminho_arquivo_limpo):
    """
    Carrega a planilha limpa gerada, realiza a análise e gera o relatório.
    """
    # Carregar a planilha limpa gerada
    df_limpo = pd.read_excel(caminho_arquivo_limpo)

    # Análise de categorias e produtos
    categorias_soma, produtos_soma = analise_categorias_produtos(df_limpo)

    # Exemplo de análise adicional que você pode adicionar
    print("\nAnálise de Categorias e Produtos finalizada.")

# **Passo 1: Gerar a planilha limpa**
# Caminho da planilha original
caminho_arquivo = 'retail_sales_data_cleaned.xlsx'
# Caminho onde a planilha limpa será salva
caminho_saida = 'retail_sales_data_cleaned.xlsx'

# Gerar a planilha limpa
carregar_e_limpar_arquivo(caminho_arquivo, caminho_saida)

# **Passo 2: Realizar a análise com a planilha limpa gerada**
realizar_analise(caminho_saida)


import pandas as pd
import numpy as np
import os

# Função para carregar arquivos de diferentes formatos (Excel, CSV, JSON, etc.)
def carregar_arquivo(file_path):
    """
    Carrega um arquivo com base na extensão para o formato adequado.
    Suporta .xls, .xlsx, .csv, .json, entre outros.
    """
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    elif file_extension == '.csv':
        df = pd.read_csv(file_path)
    elif file_extension == '.json':
        df = pd.read_json(file_path)
    elif file_extension == '.html':
        df = pd.read_html(file_path)[0]  # pd.read_html retorna uma lista, então pegamos o primeiro item
    else:
        raise ValueError(f"Formato de arquivo não suportado: {file_extension}")
    
    return df

# Função para análise detalhada dos dados
def analise_dos_dados(df):
    """
    Realiza uma análise detalhada dos dados da planilha e gera insights avançados.
    """
    print("\nIniciando análise dos dados...")

    # Exibindo os primeiros dados para entender a estrutura da planilha
    print("\nPrimeiros dados da planilha:")
    print(df.head())

    # Verificando as colunas disponíveis
    print("\nColunas disponíveis na planilha:")
    print(df.columns)

    # Analisando as categorias de produtos
    categorias = df.filter(like='Categoria_')
    categorias_soma = categorias.sum(axis=0).sort_values(ascending=False)

    # Analisando os produtos vendidos
    produtos = df.filter(like='Produto_')
    produtos_soma = produtos.sum(axis=0).sort_values(ascending=False)

    # Analisando as vendas totais
    vendas_totais = df['Valor_Total'].sum()

    # Analisando vendas por ano e mês
    df['Ano_Venda'] = pd.to_datetime(df['Data_Venda']).dt.year
    df['Mes_Venda'] = pd.to_datetime(df['Data_Venda']).dt.month
    vendas_mensais = df.groupby(['Ano_Venda', 'Mes_Venda'])['Valor_Total'].sum().reset_index()

    return categorias_soma, produtos_soma, vendas_totais, vendas_mensais

# Função para gerar o relatório detalhado para o empresário
def gerar_relatorio(df, relatorio_path):
    """
    Gera um relatório detalhado para o empresário sobre os dados da planilha e sugestões de melhoria.
    """
    with open(relatorio_path, 'w') as file:
        file.write("Relatório de Análise de Dados - Insights Profundos e Estratégias\n")
        file.write("=" * 50 + "\n")
        
        # Introdução ao relatório
        file.write("\nIntrodução:\n")
        file.write("-" * 50 + "\n")
        file.write("Este relatório foi gerado a partir dos dados fornecidos na planilha, que contém informações detalhadas sobre as vendas realizadas pela sua empresa. O objetivo deste relatório é fornecer uma análise aprofundada dos dados e sugerir ações para otimizar o desempenho de vendas, melhorar a rentabilidade e destacar oportunidades para o crescimento do seu negócio.\n")
        
        # Análise dos dados
        categorias_soma, produtos_soma, vendas_totais, vendas_mensais = analise_dos_dados(df)
        
        # Análise das Categorias
        file.write("\nAnálise das Categorias de Produtos:\n")
        file.write("-" * 50 + "\n")
        file.write("As categorias de produtos mais relevantes em termos de vendas totais são as seguintes:\n")
        file.write(f"{categorias_soma}\n")
        file.write("\nObservações e Estratégias para as Categorias:\n")
        file.write("1. Se algumas categorias estão com vendas baixas, pode ser interessante revisar estratégias de marketing ou promoções específicas para essas categorias.\n")
        file.write("2. Caso uma categoria esteja se destacando, é importante focar nela, aumentando a oferta e a promoção de produtos dessa categoria.\n")
        
        # Análise dos Produtos
        file.write("\nAnálise dos Produtos Mais Vendidos:\n")
        file.write("-" * 50 + "\n")
        file.write("Aqui estão os produtos mais vendidos com maior impacto nas vendas totais:\n")
        file.write(f"{produtos_soma}\n")
        file.write("\nSugestões para os Produtos:\n")
        file.write("1. Produtos com alta demanda devem ter um aumento de estoque ou campanhas de marketing direcionadas.\n")
        file.write("2. Produtos com baixo desempenho devem ser analisados para entender se o preço, promoção ou demanda está impactando suas vendas.\n")
        
        # Análise das Vendas Totais
        file.write("\nAnálise de Vendas Totais:\n")
        file.write("-" * 50 + "\n")
        file.write(f"Vendas totais realizadas: R${vendas_totais:.2f}\n")
        file.write("\nEstratégias para Melhorar as Vendas Totais:\n")
        file.write("1. Acompanhe a performance de vendas de forma mensal para identificar padrões sazonais e preparar promoções.\n")
        file.write("2. Se houver meses com vendas muito baixas, considerar a introdução de novos produtos ou estratégias de descontos.\n")
        
        # Vendas por Ano e Mês
        file.write("\nAnálise de Vendas ao Longo do Tempo (Ano/Mês):\n")
        file.write("-" * 50 + "\n")
        file.write(f"{vendas_mensais}\n")
        file.write("\nSugestões para Vendas ao Longo do Tempo:\n")
        file.write("1. Avalie a sazonalidade das vendas. Se houver meses com vendas mais baixas, considere campanhas de marketing ou descontos para atrair mais clientes.\n")
        file.write("2. Com base nas tendências mensais, planeje ações para aumentar as vendas em meses de baixa performance.\n")

        # Conclusão
        file.write("\nConclusão:\n")
        file.write("-" * 50 + "\n")
        file.write("Com base na análise dos dados, recomendamos que sua empresa foque em promover as categorias e produtos com melhor desempenho, ao mesmo tempo que revisa as estratégias de marketing para categorias e produtos com baixo desempenho. Acompanhando as vendas ao longo do tempo, é possível identificar padrões e ajustar as campanhas promocionais de forma estratégica para melhorar os resultados.\n")

# Função para carregar, limpar e salvar o arquivo
def carregar_e_limpar_arquivo(caminho_arquivo, caminho_saida, relatorio_path):
    """
    Carrega, limpa os dados, gera o relatório e salva a planilha limpa.
    """
    # Carregar o arquivo com base no formato
    df = carregar_arquivo(caminho_arquivo)

    # Gerar relatório com base nos dados
    gerar_relatorio(df, relatorio_path)
    print(f"Relatório gerado e salvo em: {relatorio_path}")

    # Salvar a planilha limpa (aplicando as transformações)
    df.to_excel(caminho_saida, index=False)
    print(f"Planilha limpa salva em: {caminho_saida}")

# Caminhos dos arquivos


caminho_entrada = 'retail_sales_data.xlsx'
caminho_saida = 'retail_sales_data.xlsx'
relatorio_path = r'C:\Users\mycha\OneDrive\Área de Trabalho\teste\Projetos em análise de dados\Sistema de Análise e Previsão de Demanda\relatorio_analise.txt'


# Chama a função para carregar, limpar e gerar o relatório
carregar_e_limpar_arquivo(caminho_entrada, caminho_saida, relatorio_path)

