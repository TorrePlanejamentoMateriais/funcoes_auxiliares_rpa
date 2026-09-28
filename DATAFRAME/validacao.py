"""
============================================================
Módulo: DATAFRAME / validacao.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 24/09/2026
Última Alteração: 24/09/2026
Versão: 1.0.0

Descrição:
    Biblioteca de funções auxiliares para validação,
    análise e inspeção de DataFrames utilizando pandas.

Funções Disponíveis:
    - verifica_colunas()
    - verifica_tipos_colunas()
    - verifica_nulos()
    - verifica_duplicados()

Dependências:
    - pandas

Histórico:
    v1.0.0 - 24/09/2026
        - Criação inicial do módulo.
        - Implementação das funções de validação.
============================================================
"""

# ----------------------------------------------------------------------------
# --------------------------- IMPORTS ----------------------------------------
# ----------------------------------------------------------------------------

# Importa a biblioteca pandas utilizando o apelido pd
import pandas as pd


# ----------------------------------------------------------------------------
# --------------------------- FUNÇÕES ----------------------------------------
# ----------------------------------------------------------------------------

def verifica_colunas(dataframe: pd.DataFrame, colunas: list[str]):
    """
    Verifica quais colunas de uma lista estão presentes em um DataFrame.

    Percorre as colunas informadas e separa os nomes entre colunas
    encontradas e colunas não encontradas no DataFrame.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém as colunas que serão verificadas.

        colunas (list[str]):
            Lista com os nomes das colunas que devem ser verificadas.

    Returns:
        dict:
            Dicionário contendo as colunas encontradas e não encontradas,
            juntamente com a quantidade de colunas de cada grupo.

    Raises:
        Exception:
            Repassa qualquer erro ocorrido durante a verificação das colunas.
    """

    # Cria uma lista vazia para armazenar as colunas encontradas no DataFrame
    colunas_existentes = []

    # Cria uma lista vazia para armazenar as colunas não encontradas no DataFrame
    colunas_faltantes = []

    # Inicia o bloco de tratamento de exceções
    try:

        # Percorre cada coluna informada na lista de colunas
        for coluna in colunas:

            # Verifica se a coluna atual existe entre as colunas do DataFrame
            if coluna in dataframe.columns:

                # Adiciona a coluna atual à lista de colunas existentes
                colunas_existentes.append(coluna)

            # Executa este bloco quando a coluna atual não existe no DataFrame
            else:

                # Adiciona a coluna atual à lista de colunas faltantes
                colunas_faltantes.append(coluna)

        # Cria o dicionário com as colunas encontradas e não encontradas
        resultado = {
            'Colunas Encontradas': {
                'Colunas': colunas_existentes,
                'Quantidade': len(colunas_existentes)
            },
            'Colunas Nao Encontradas': {
                'Colunas': colunas_faltantes,
                'Quantidade': len(colunas_faltantes)
            }
        }

        # Retorna o resultado da verificação
        return resultado

    # Captura qualquer exceção ocorrida dentro do bloco try
    except Exception as erro:

        # Exibe uma mensagem informando o erro ocorrido
        print(f'Erro ao verificar as colunas no dataframe: {erro}')

        # Lança novamente a exceção capturada, preservando o fluxo de erro
        raise


def verifica_tipos_colunas(dataframe: pd.DataFrame):
    """
    Identifica os tipos de dados das colunas de um DataFrame.

    Retorna a quantidade total de colunas e um dicionário contendo
    o nome e o tipo de dado de cada coluna.

    Args:
        dataframe (pd.DataFrame):
            DataFrame cujos tipos das colunas serão identificados.

    Returns:
        dict:
            Dicionário contendo a quantidade total de colunas e os
            respectivos tipos de dados.

            Caso o DataFrame esteja vazio, retorna um dicionário vazio.

    Raises:
        TypeError:
            Ocorre quando o objeto fornecido não é um DataFrame do pandas.

        Exception:
            Ocorre quando não é possível processar os tipos das colunas.
    """

    # Cria um dicionário vazio para armazenar o resultado da verificação
    resultado = {}

    # Inicia o bloco de tratamento de exceções
    try:

        # Verifica se o objeto recebido é uma instância de DataFrame do pandas
        if isinstance(dataframe, pd.DataFrame):

            # Verifica se o DataFrame não possui registros
            if dataframe.empty:

                # Exibe uma mensagem informando que o DataFrame está vazio
                print(f'O DataFrame está vazio')

                # Retorna o dicionário vazio
                return resultado

            # Cria um dicionário com a quantidade de colunas e o tipo de cada coluna
            resultado = {
                'Quantidade Colunas': len(dataframe.columns),
                'Colunas': {
                    coluna: str(tipo)
                    for coluna, tipo in dataframe.dtypes.items()
                }
            }

            # Retorna o resultado contendo os tipos das colunas
            return resultado

        # Executa este bloco caso o objeto recebido não seja um DataFrame
        else:

            # Gera uma exceção informando que o objeto não é um DataFrame
            raise TypeError("O objeto fornecido não é um DataFrame do pandas.")

    # Captura qualquer exceção ocorrida durante o processamento
    except Exception as e:

        # Gera uma nova exceção contendo a descrição do erro original
        raise Exception(f"Erro ao processar DataFrame: {e}")


def verifica_nulos(dataframe: pd.DataFrame):
    """
    Verifica a quantidade de valores nulos em cada coluna de um DataFrame.

    A função analisa todas as colunas e contabiliza os valores nulos,
    como None, NaN e NaT, presentes em cada uma delas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será analisado.

    Returns:
        dict | None:
            Dicionário em que cada chave representa o nome de uma coluna
            e cada valor representa sua quantidade de valores nulos.

            Retorna None quando o DataFrame está vazio.

    Observações:
        Caso o objeto fornecido não seja um DataFrame ou ocorra algum erro
        durante o processamento, uma mensagem será exibida no terminal.
    """

    # Inicia o bloco de tratamento de exceções
    try:

        # Verifica se o objeto recebido é uma instância de DataFrame do pandas
        if isinstance(dataframe, pd.DataFrame):

            # Verifica se o DataFrame não possui registros
            if dataframe.empty:

                # Exibe uma mensagem informando que o DataFrame está vazio
                print("O DataFrame está vazio.")

                # Encerra a função retornando None
                return None

            # Identifica os valores nulos, soma por coluna e converte o resultado em um dicionário
            resultado = dataframe.isna().sum().to_dict()

        # Executa este bloco caso o objeto recebido não seja um DataFrame
        else:

            # Exibe uma mensagem informando que a variável não é um DataFrame
            print("A variável fornecida não é um DataFrame.")

    # Captura qualquer exceção ocorrida durante o processamento
    except Exception as e:

        # Exibe uma mensagem contendo a descrição do erro
        print(f"Erro ao processar: {e}")

    # Retorna o dicionário com a quantidade de valores nulos de cada coluna
    return resultado


def verifica_duplicados(dataframe: pd.DataFrame):
    """
    Verifica a quantidade de valores duplicados em cada coluna de um DataFrame.

    A função analisa cada coluna individualmente e contabiliza os valores
    que aparecem novamente após sua primeira ocorrência.

    Args:
        dataframe (pd.DataFrame):
            DataFrame cujas colunas serão verificadas.

    Returns:
        dict | None:
            Dicionário em que cada chave representa o nome de uma coluna
            e cada valor representa sua quantidade de valores duplicados.

            Retorna None quando o DataFrame está vazio.

    Observações:
        Caso o objeto fornecido não seja um DataFrame ou ocorra algum erro
        durante o processamento, uma mensagem será exibida no terminal.
    """
    
    # Inicia o bloco de tratamento de exceções
    try:

        # Verifica se o objeto recebido é uma instância de DataFrame do pandas
        if isinstance(dataframe, pd.DataFrame):

            # Verifica se o DataFrame não possui registros
            if dataframe.empty:

                # Exibe uma mensagem informando que o DataFrame está vazio
                print("O DataFrame está vazio.")

                # Encerra a função retornando None
                return None

            # Aplica a verificação em cada coluna, soma os valores duplicados e converte o resultado em dicionário
            resultado = dataframe.apply(lambda col: col.duplicated().sum()).to_dict()

        # Executa este bloco caso o objeto recebido não seja um DataFrame
        else:

            # Exibe uma mensagem informando que a variável não é um DataFrame
            print("A variável fornecida não é um DataFrame.")

    # Captura qualquer exceção ocorrida durante o processamento
    except Exception as e:

        # Exibe uma mensagem contendo a descrição do erro
        print(f"Erro ao processar: {e}")

    # Retorna o dicionário com a quantidade de valores duplicados de cada coluna
    return resultado