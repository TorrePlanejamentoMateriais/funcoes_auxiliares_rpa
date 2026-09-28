"""
============================================================
Módulo: DATAFRAME / colunas.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 25/09/2026
Última Alteração: 25/09/2026
Versão: 1.0.0

Descrição:
    Biblioteca de funções auxiliares para manipulação
    das colunas de DataFrames utilizando pandas.

Funções disponíveis:
    - renomear_colunas()
    - selecionar_colunas()
    - remover_colunas()

Dependências:
    - pandas

Histórico:
    v1.0.0 - 25/09/2026
        - Criação inicial do módulo.
        - Implementação da função renomear_colunas.
        - Implementação da função selecionar_colunas.
        - Implementação da função remover_colunas.
        - Inclusão de validações de tipos e colunas.
        - Garantia de preservação do DataFrame original.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa a biblioteca pandas utilizando o apelido pd
import pandas as pd


# ----------------------------------------------------------------------------
# FUNÇÕES
# ----------------------------------------------------------------------------


def renomear_colunas(
    dataframe: pd.DataFrame,
    colunas: list[str],
    novos_nomes: list[str]
) -> pd.DataFrame:
    """
    Renomeia colunas específicas de um DataFrame.

    A função cria uma cópia profunda do DataFrame original e renomeia
    somente as colunas informadas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá as colunas renomeadas.

        colunas (list[str]):
            Lista contendo os nomes atuais das colunas que serão
            renomeadas.

        novos_nomes (list[str]):
            Lista contendo os novos nomes que serão atribuídos
            às colunas.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com as colunas renomeadas.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro colunas não seja uma lista.

        TypeError:
            Caso o parâmetro novos_nomes não seja uma lista.

        TypeError:
            Caso algum item das listas não seja uma string.

        ValueError:
            Caso alguma das listas esteja vazia.

        ValueError:
            Caso as listas não possuam a mesma quantidade de elementos.

        ValueError:
            Caso a lista de colunas possua nomes duplicados.

        ValueError:
            Caso os novos nomes possuam valores vazios ou duplicados.

        ValueError:
            Caso a renomeação gere colunas duplicadas no DataFrame.

        KeyError:
            Caso alguma coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o parâmetro colunas é uma lista
    if not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings."
        )

    # Verifica se o parâmetro novos_nomes é uma lista
    if not isinstance(novos_nomes, list):
        raise TypeError(
            "O parâmetro 'novos_nomes' deve ser uma lista de strings."
        )

    # Verifica se a lista de colunas está vazia
    if not colunas:
        raise ValueError(
            "A lista 'colunas' não pode estar vazia."
        )

    # Verifica se a lista de novos nomes está vazia
    if not novos_nomes:
        raise ValueError(
            "A lista 'novos_nomes' não pode estar vazia."
        )

    # Verifica se todos os nomes atuais são strings
    if not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Verifica se todos os novos nomes são strings
    if not all(
        isinstance(novo_nome, str)
        for novo_nome in novos_nomes
    ):
        raise TypeError(
            "Todos os itens da lista 'novos_nomes' devem ser strings."
        )

    # Verifica se as listas possuem a mesma quantidade de elementos
    if len(colunas) != len(novos_nomes):
        raise ValueError(
            "A quantidade de colunas deve ser igual à quantidade "
            "de novos nomes."
        )

    # Verifica se existem nomes vazios na lista de novos nomes
    novos_nomes_vazios = [
        novo_nome
        for novo_nome in novos_nomes
        if not novo_nome.strip()
    ]

    # Interrompe a execução quando algum novo nome está vazio
    if novos_nomes_vazios:
        raise ValueError(
            "A lista 'novos_nomes' não pode possuir nomes vazios."
        )

    # Localiza nomes repetidos na lista de colunas atuais
    colunas_duplicadas = {
        coluna
        for coluna in colunas
        if colunas.count(coluna) > 1
    }

    # Interrompe a execução quando existem colunas repetidas
    if colunas_duplicadas:
        raise ValueError(
            "A lista 'colunas' possui nomes duplicados: "
            f"{sorted(colunas_duplicadas)}"
        )

    # Localiza as colunas informadas que não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução quando alguma coluna não é encontrada
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Localiza nomes repetidos na lista de novos nomes
    novos_nomes_duplicados = {
        novo_nome
        for novo_nome in novos_nomes
        if novos_nomes.count(novo_nome) > 1
    }

    # Interrompe a execução quando existem novos nomes repetidos
    if novos_nomes_duplicados:
        raise ValueError(
            "A lista 'novos_nomes' possui nomes duplicados: "
            f"{sorted(novos_nomes_duplicados)}"
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    # Cria o mapeamento entre os nomes atuais e os novos nomes
    mapeamento_colunas = dict(
        zip(colunas, novos_nomes)
    )

    # Renomeia somente as colunas especificadas
    dataframe_tratado = dataframe_tratado.rename(
        columns=mapeamento_colunas
    )

    # Localiza nomes duplicados após a renomeação
    colunas_resultantes_duplicadas = (
        dataframe_tratado.columns[
            dataframe_tratado.columns.duplicated()
        ]
        .unique()
        .tolist()
    )

    # Interrompe a execução quando a renomeação gera nomes duplicados
    if colunas_resultantes_duplicadas:
        raise ValueError(
            "A renomeação gerou colunas duplicadas no DataFrame: "
            f"{colunas_resultantes_duplicadas}"
        )

    # Retorna a cópia do DataFrame com as colunas renomeadas
    return dataframe_tratado


def selecionar_colunas(
    dataframe: pd.DataFrame,
    colunas: list[str]
) -> pd.DataFrame:
    """
    Seleciona colunas específicas de um DataFrame.

    A função cria uma cópia profunda contendo somente as colunas
    informadas, mantendo a ordem definida na lista.

    Args:
        dataframe (pd.DataFrame):
            DataFrame do qual as colunas serão selecionadas.

        colunas (list[str]):
            Lista contendo os nomes das colunas que deverão permanecer
            no DataFrame retornado.

    Returns:
        pd.DataFrame:
            Novo DataFrame contendo somente as colunas selecionadas.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro colunas não seja uma lista.

        TypeError:
            Caso algum item da lista de colunas não seja uma string.

        ValueError:
            Caso a lista de colunas esteja vazia.

        ValueError:
            Caso a lista possua nomes de colunas duplicados.

        KeyError:
            Caso alguma coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o parâmetro colunas é uma lista
    if not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings."
        )

    # Verifica se a lista de colunas está vazia
    if not colunas:
        raise ValueError(
            "A lista 'colunas' não pode estar vazia."
        )

    # Verifica se todos os nomes informados são strings
    if not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Localiza nomes repetidos na lista de colunas
    colunas_duplicadas = {
        coluna
        for coluna in colunas
        if colunas.count(coluna) > 1
    }

    # Interrompe a execução quando existem colunas repetidas
    if colunas_duplicadas:
        raise ValueError(
            "A lista 'colunas' possui nomes duplicados: "
            f"{sorted(colunas_duplicadas)}"
        )

    # Localiza as colunas informadas que não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução quando alguma coluna não é encontrada
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Seleciona as colunas e cria uma cópia profunda do resultado
    dataframe_tratado = dataframe.loc[
        :,
        colunas
    ].copy(deep=True)

    # Retorna o DataFrame contendo somente as colunas selecionadas
    return dataframe_tratado


def remover_colunas(
    dataframe: pd.DataFrame,
    colunas: list[str]
) -> pd.DataFrame:
    """
    Remove colunas específicas de um DataFrame.

    A função cria uma cópia profunda do DataFrame original e remove
    somente as colunas informadas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame do qual as colunas serão removidas.

        colunas (list[str]):
            Lista contendo os nomes das colunas que serão removidas.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame sem as colunas removidas.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro colunas não seja uma lista.

        TypeError:
            Caso algum item da lista de colunas não seja uma string.

        ValueError:
            Caso a lista de colunas esteja vazia.

        ValueError:
            Caso a lista possua nomes de colunas duplicados.

        ValueError:
            Caso a remoção deixe o DataFrame sem nenhuma coluna.

        KeyError:
            Caso alguma coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o parâmetro colunas é uma lista
    if not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings."
        )

    # Verifica se a lista de colunas está vazia
    if not colunas:
        raise ValueError(
            "A lista 'colunas' não pode estar vazia."
        )

    # Verifica se todos os nomes informados são strings
    if not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Localiza nomes repetidos na lista de colunas
    colunas_duplicadas = {
        coluna
        for coluna in colunas
        if colunas.count(coluna) > 1
    }

    # Interrompe a execução quando existem colunas repetidas
    if colunas_duplicadas:
        raise ValueError(
            "A lista 'colunas' possui nomes duplicados: "
            f"{sorted(colunas_duplicadas)}"
        )

    # Localiza as colunas informadas que não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução quando alguma coluna não é encontrada
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Verifica se todas as colunas do DataFrame seriam removidas
    if len(colunas) == len(dataframe.columns):
        raise ValueError(
            "Não é possível remover todas as colunas do DataFrame."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    # Remove as colunas selecionadas e armazena o resultado
    dataframe_tratado = dataframe_tratado.drop(
        columns=colunas
    )

    # Retorna a cópia do DataFrame sem as colunas removidas
    return dataframe_tratado