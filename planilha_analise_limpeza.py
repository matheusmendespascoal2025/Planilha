import pandas as pd
import numpy as np

def carregar_planilha(caminho_arquivo):
    """
    Carrega dados de um arquivo CSV ou Excel (.xlsx).
    Detecta automaticamente o formato do arquivo.
    """
    try:
        if caminho_arquivo.endswith('.csv'):
            print(f"Carregando CSV: {caminho_arquivo}")
            df = pd.read_csv(caminho_arquivo)
        elif caminho_arquivo.endswith('.xlsx'):
            print(f"Carregando Excel: {caminho_arquivo}")
            df = pd.read_excel(caminho_arquivo)
        else:
            raise ValueError("Formato de arquivo não suportado. Use .csv ou .xlsx.")
        print("Planilha carregada com sucesso!")
        return df
    except FileNotFoundError:
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro ao carregar a planilha: {e}")
        return None

def limpar_planilha(df):
    """
    Realiza operações básicas de limpeza de dados:
    - Remove linhas completamente duplicadas.
    - Preenche valores ausentes em colunas numéricas com a mediana.
    - Preenche valores ausentes em colunas categóricas com o modo (valor mais frequente).
    """
    if df is None:
        return None

    print("\n--- Iniciando Limpeza de Dados ---")

    # 1. Remover linhas duplicadas
    linhas_antes = len(df)
    df.drop_duplicates(inplace=True)
    linhas_depois = len(df)
    if linhas_antes > linhas_depois:
        print(f"Removidas {linhas_antes - linhas_depois} linhas duplicadas.")
    else:
        print("Nenhuma linha duplicada encontrada.")

    # 2. Preencher valores ausentes
    print("Verificando e preenchendo valores ausentes...")
    for coluna in df.columns:
        if df[coluna].isnull().any():
            if pd.api.types.is_numeric_dtype(df[coluna]):
                mediana = df[coluna].median()
                df[coluna].fillna(mediana, inplace=True)
                print(f"  Coluna '{coluna}': Valores ausentes preenchidos com a mediana ({mediana}).")
            elif pd.api.types.is_string_dtype(df[coluna]) or pd.api.types.is_object_dtype(df[coluna]):
                modo = df[coluna].mode()[0] if not df[coluna].mode().empty else 'N/A'
                df[coluna].fillna(modo, inplace=True)
                print(f"  Coluna '{coluna}': Valores ausentes preenchidos com o modo ('{modo}').")
            else:
                df[coluna].fillna('Desconhecido', inplace=True) # Para outros tipos
                print(f"  Coluna '{coluna}': Valores ausentes preenchidos com 'Desconhecido'.")
        else:
            print(f"  Coluna '{coluna}': Sem valores ausentes.")

    print("Limpeza de dados concluída!")
    return df

def analisar_planilha(df):
    """
    Apresenta uma análise básica do DataFrame.
    """
    if df is None:
        return None

    print("\n--- Análise Básica da Planilha ---")
    print("\nInformações gerais do DataFrame:")
    df.info()

    print("\nEstatísticas Descritivas (para colunas numéricas):")
    print(df.describe())

    print("\nContagem de valores únicos nas primeiras 5 colunas (para colunas categóricas):")
    for col in df.columns[:5]: # Limita às 5 primeiras para evitar saída muito longa
        if not pd.api.types.is_numeric_dtype(df[col]):
            print(f"  Coluna '{col}':\n{df[col].value_counts().head(5)}") # Mostra os 5 mais frequentes
            print("-" * 20)
    return df

def salvar_planilha(df, caminho_saida, formato='csv'):
    """
    Salva o DataFrame processado em um novo arquivo.
    """
    if df is None:
        print("Nenhum DataFrame para salvar.")
        return

    try:
        if formato.lower() == 'csv':
            df.to_csv(caminho_saida, index=False)
        elif formato.lower() == 'xlsx':
            df.to_excel(caminho_saida, index=False)
        else:
            raise ValueError("Formato de saída não suportado. Use 'csv' ou 'xlsx'.")
        print(f"\nPlanilha processada salva em: {caminho_saida}")
    except Exception as e:
        print(f"Ocorreu um erro ao salvar a planilha: {e}")

if __name__ == "__main__":
    # --- Configuração ---
    # Altere este caminho para o seu arquivo de planilha
    # Exemplo: 'dados_vendas.csv' ou 'relatorio_financeiro.xlsx'
    arquivo_entrada = 'seus_dados.csv' # Ou 'seus_dados.xlsx'
    arquivo_saida = 'seus_dados_limpos.csv' # Ou 'seus_dados_limpos.xlsx'
    formato_saida = 'csv' # Ou 'xlsx'

    # --- Fluxo Principal ---
    df_original = carregar_planilha(arquivo_entrada)

    if df_original is not None:
        df_limpo = limpar_planilha(df_original.copy()) # Usa uma cópia para não alterar o original
        df_final = analisar_planilha(df_limpo)
        salvar_planilha(df_final, arquivo_saida, formato=formato_saida)
    else:
        print("Nenhum dado processado devido a erro no carregamento.")

    print("\nProcessamento da planilha concluído!")
