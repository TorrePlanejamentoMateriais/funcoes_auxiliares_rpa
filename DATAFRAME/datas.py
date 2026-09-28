"""
============================================================
Módulo: DATAFRAME / datas.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 25/09/2026
Última Alteração: 25/09/2026
Versão: 1.0.0

Descrição:
    Biblioteca de funções auxiliares para extração e cálculo
    de informações relacionadas a datas em DataFrames
    utilizando pandas.

Funções disponíveis:
    - extrair_ano()
    - extrair_mes()
    - extrair_dia()
    - dias_entre_datas()
    - coluna_atrasada()

Dependências:
    - pandas

Histórico:
    v1.0.0 - 25/09/2026
        - Criação inicial do módulo.
        - Implementação da função extrair_ano.
        - Implementação da função extrair_mes.
        - Implementação da função extrair_dia.
        - Implementação da função dias_entre_datas.
        - Implementação da função coluna_atrasada.
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


def extrair_ano(
    dataframe: pd.DataFrame,
    coluna: str
) -> pd.DataFrame:
    """
    Extrai o ano de uma coluna de data.

    A função cria uma cópia profunda do DataFrame original,
    converte a coluna informada para datetime e cria uma nova
    coluna chamada ANO.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém a coluna de data.

        coluna (str):
            Nome da coluna da qual o ano será extraído.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com a nova coluna ANO.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro coluna não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso o parâmetro coluna esteja vazio.

        ValueError:
            Caso algum valor não possa ser convertido para data.

        KeyError:
            Caso a coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame está vazio
    if dataframe.empty:
        raise ValueError(
            "O DataFrame está vazio."
        )

    # Verifica se o parâmetro coluna é uma string
    if not isinstance(coluna, str):
        raise TypeError(
            "O parâmetro 'coluna' deve ser uma string."
        )

    # Verifica se o nome da coluna está vazio
    if not coluna.strip():
        raise ValueError(
            "O parâmetro 'coluna' não pode estar vazio."
        )

    # Verifica se a coluna existe no DataFrame
    if coluna not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna}' não foi encontrada no DataFrame."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    try:

        # Converte a coluna informada para o tipo datetime
        coluna_convertida = pd.to_datetime(
            dataframe_tratado[coluna],
            errors="raise",
            dayfirst=True
        )

    except (ValueError, TypeError) as erro:
        raise ValueError(
            f"Não foi possível converter a coluna '{coluna}' para data. "
            f"Erro original: {erro}"
        ) from erro

    # Cria a coluna ANO utilizando o ano da data convertida
    dataframe_tratado["ANO"] = (
        coluna_convertida.dt.year.astype("Int64")
    )

    # Retorna a cópia do DataFrame com a coluna ANO
    return dataframe_tratado


def extrair_mes(
    dataframe: pd.DataFrame,
    coluna: str
) -> pd.DataFrame:
    """
    Extrai o mês de uma coluna de data.

    A função cria uma cópia profunda do DataFrame original,
    converte a coluna informada para datetime e cria uma nova
    coluna chamada MES, contendo o mês por extenso.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém a coluna de data.

        coluna (str):
            Nome da coluna da qual o mês será extraído.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com a nova coluna MES.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro coluna não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso o parâmetro coluna esteja vazio.

        ValueError:
            Caso algum valor não possa ser convertido para data.

        KeyError:
            Caso a coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame está vazio
    if dataframe.empty:
        raise ValueError(
            "O DataFrame está vazio."
        )

    # Verifica se o parâmetro coluna é uma string
    if not isinstance(coluna, str):
        raise TypeError(
            "O parâmetro 'coluna' deve ser uma string."
        )

    # Verifica se o nome da coluna está vazio
    if not coluna.strip():
        raise ValueError(
            "O parâmetro 'coluna' não pode estar vazio."
        )

    # Verifica se a coluna existe no DataFrame
    if coluna not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna}' não foi encontrada no DataFrame."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    # Cria um dicionário contendo os meses por extenso
    meses_por_extenso = {
        1: "JANEIRO",
        2: "FEVEREIRO",
        3: "MARÇO",
        4: "ABRIL",
        5: "MAIO",
        6: "JUNHO",
        7: "JULHO",
        8: "AGOSTO",
        9: "SETEMBRO",
        10: "OUTUBRO",
        11: "NOVEMBRO",
        12: "DEZEMBRO"
    }

    try:

        # Converte a coluna informada para o tipo datetime
        coluna_convertida = pd.to_datetime(
            dataframe_tratado[coluna],
            errors="raise",
            dayfirst=True
        )

    except (ValueError, TypeError) as erro:
        raise ValueError(
            f"Não foi possível converter a coluna '{coluna}' para data. "
            f"Erro original: {erro}"
        ) from erro

    # Extrai o número do mês da coluna convertida
    numero_mes = coluna_convertida.dt.month.astype("Int64")

    # Cria a coluna MES utilizando o nome do mês por extenso
    dataframe_tratado["MES"] = numero_mes.map(
        meses_por_extenso
    )

    # Retorna a cópia do DataFrame com a coluna MES
    return dataframe_tratado


def extrair_dia(
    dataframe: pd.DataFrame,
    coluna: str
) -> pd.DataFrame:
    """
    Extrai o dia do mês de uma coluna de data.

    A função cria uma cópia profunda do DataFrame original,
    converte a coluna informada para datetime e cria uma nova
    coluna chamada DIA.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém a coluna de data.

        coluna (str):
            Nome da coluna da qual o dia será extraído.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com a nova coluna DIA.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro coluna não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso o parâmetro coluna esteja vazio.

        ValueError:
            Caso algum valor não possa ser convertido para data.

        KeyError:
            Caso a coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame está vazio
    if dataframe.empty:
        raise ValueError(
            "O DataFrame está vazio."
        )

    # Verifica se o parâmetro coluna é uma string
    if not isinstance(coluna, str):
        raise TypeError(
            "O parâmetro 'coluna' deve ser uma string."
        )

    # Verifica se o nome da coluna está vazio
    if not coluna.strip():
        raise ValueError(
            "O parâmetro 'coluna' não pode estar vazio."
        )

    # Verifica se a coluna existe no DataFrame
    if coluna not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna}' não foi encontrada no DataFrame."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    try:

        # Converte a coluna informada para o tipo datetime
        coluna_convertida = pd.to_datetime(
            dataframe_tratado[coluna],
            errors="raise",
            dayfirst=True
        )

    except (ValueError, TypeError) as erro:
        raise ValueError(
            f"Não foi possível converter a coluna '{coluna}' para data. "
            f"Erro original: {erro}"
        ) from erro

    # Cria a coluna DIA utilizando o dia da data convertida
    dataframe_tratado["DIA"] = (
        coluna_convertida.dt.day.astype("Int64")
    )

    # Retorna a cópia do DataFrame com a coluna DIA
    return dataframe_tratado


def dias_entre_datas(
    dataframe: pd.DataFrame,
    coluna_data_inicial: str,
    coluna_data_final: str
) -> pd.DataFrame:
    """
    Calcula a quantidade de dias entre duas colunas de datas.

    A função realiza o cálculo:

        data final - data inicial

    O resultado poderá ser negativo quando a data final for
    anterior à data inicial.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém as colunas de datas.

        coluna_data_inicial (str):
            Nome da coluna que contém a data inicial.

        coluna_data_final (str):
            Nome da coluna que contém a data final.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com a nova coluna DIAS_ENTRE_DATAS.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso algum nome de coluna não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso algum nome de coluna esteja vazio.

        ValueError:
            Caso as duas colunas informadas sejam iguais.

        ValueError:
            Caso algum valor não possa ser convertido para data.

        KeyError:
            Caso alguma das colunas não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame está vazio
    if dataframe.empty:
        raise ValueError(
            "O DataFrame está vazio."
        )

    # Verifica se a coluna inicial é uma string
    if not isinstance(coluna_data_inicial, str):
        raise TypeError(
            "O parâmetro 'coluna_data_inicial' deve ser uma string."
        )

    # Verifica se a coluna final é uma string
    if not isinstance(coluna_data_final, str):
        raise TypeError(
            "O parâmetro 'coluna_data_final' deve ser uma string."
        )

    # Verifica se o nome da coluna inicial está vazio
    if not coluna_data_inicial.strip():
        raise ValueError(
            "O parâmetro 'coluna_data_inicial' não pode estar vazio."
        )

    # Verifica se o nome da coluna final está vazio
    if not coluna_data_final.strip():
        raise ValueError(
            "O parâmetro 'coluna_data_final' não pode estar vazio."
        )

    # Verifica se a coluna inicial existe no DataFrame
    if coluna_data_inicial not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna_data_inicial}' não foi encontrada "
            "no DataFrame."
        )

    # Verifica se a coluna final existe no DataFrame
    if coluna_data_final not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna_data_final}' não foi encontrada "
            "no DataFrame."
        )

    # Verifica se as colunas informadas são diferentes
    if coluna_data_inicial == coluna_data_final:
        raise ValueError(
            "As colunas de data inicial e data final devem ser diferentes."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    try:

        # Converte a coluna de data inicial para datetime
        data_inicial = pd.to_datetime(
            dataframe_tratado[coluna_data_inicial],
            errors="raise",
            dayfirst=True
        )

        # Converte a coluna de data final para datetime
        data_final = pd.to_datetime(
            dataframe_tratado[coluna_data_final],
            errors="raise",
            dayfirst=True
        )

    except (ValueError, TypeError) as erro:
        raise ValueError(
            "Não foi possível converter uma das colunas para data. "
            f"Erro original: {erro}"
        ) from erro

    # Calcula a diferença entre a data final e a data inicial
    diferenca_dias = (
        data_final - data_inicial
    ).dt.days.astype("Int64")

    # Cria a coluna contendo a diferença em dias
    dataframe_tratado["DIAS_ENTRE_DATAS"] = diferenca_dias

    # Retorna a cópia do DataFrame com a diferença em dias
    return dataframe_tratado


def coluna_atrasada(
    dataframe: pd.DataFrame,
    coluna_data_limite: str,
    coluna_data_conclusao: str
) -> pd.DataFrame:
    """
    Identifica se um registro foi concluído com atraso.

    Um registro será considerado atrasado quando a data de
    conclusão for posterior à data limite.

    A nova coluna ATRASADA receberá:

    - SIM quando a conclusão ocorrer depois da data limite.
    - NÃO quando a conclusão ocorrer até a data limite.
    - Valor ausente quando uma das datas estiver vazia.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém as colunas de datas.

        coluna_data_limite (str):
            Nome da coluna que contém a data limite.

        coluna_data_conclusao (str):
            Nome da coluna que contém a data de conclusão.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com a nova coluna ATRASADA.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso algum nome de coluna não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso algum nome de coluna esteja vazio.

        ValueError:
            Caso as duas colunas informadas sejam iguais.

        ValueError:
            Caso algum valor preenchido não possa ser convertido
            para data.

        KeyError:
            Caso alguma coluna informada não exista no DataFrame.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame está vazio
    if dataframe.empty:
        raise ValueError(
            "O DataFrame está vazio."
        )

    # Verifica se a coluna de data limite é uma string
    if not isinstance(coluna_data_limite, str):
        raise TypeError(
            "O parâmetro 'coluna_data_limite' deve ser uma string."
        )

    # Verifica se a coluna de conclusão é uma string
    if not isinstance(coluna_data_conclusao, str):
        raise TypeError(
            "O parâmetro 'coluna_data_conclusao' deve ser uma string."
        )

    # Verifica se o nome da coluna de data limite está vazio
    if not coluna_data_limite.strip():
        raise ValueError(
            "O parâmetro 'coluna_data_limite' não pode estar vazio."
        )

    # Verifica se o nome da coluna de conclusão está vazio
    if not coluna_data_conclusao.strip():
        raise ValueError(
            "O parâmetro 'coluna_data_conclusao' não pode estar vazio."
        )

    # Verifica se a coluna de data limite existe no DataFrame
    if coluna_data_limite not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna_data_limite}' não foi encontrada "
            "no DataFrame."
        )

    # Verifica se a coluna de conclusão existe no DataFrame
    if coluna_data_conclusao not in dataframe.columns:
        raise KeyError(
            f"A coluna '{coluna_data_conclusao}' não foi encontrada "
            "no DataFrame."
        )

    # Verifica se as colunas informadas são diferentes
    if coluna_data_limite == coluna_data_conclusao:
        raise ValueError(
            "As colunas de data limite e conclusão devem ser diferentes."
        )

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    # Identifica os valores preenchidos na coluna de data limite
    valores_limite_preenchidos = (
        dataframe_tratado[coluna_data_limite].notna()
    )

    # Identifica os valores preenchidos na coluna de conclusão
    valores_conclusao_preenchidos = (
        dataframe_tratado[coluna_data_conclusao].notna()
    )

    # Converte a coluna da data limite para datetime
    data_limite = pd.to_datetime(
        dataframe_tratado[coluna_data_limite],
        errors="coerce",
        dayfirst=True
    )

    # Converte a coluna da data de conclusão para datetime
    data_conclusao = pd.to_datetime(
        dataframe_tratado[coluna_data_conclusao],
        errors="coerce",
        dayfirst=True
    )

    # Identifica valores preenchidos que não foram convertidos
    datas_limite_invalidas = (
        valores_limite_preenchidos
        & data_limite.isna()
    )

    # Identifica conclusões preenchidas que não foram convertidas
    datas_conclusao_invalidas = (
        valores_conclusao_preenchidos
        & data_conclusao.isna()
    )

    # Interrompe a execução quando existem datas limite inválidas
    if datas_limite_invalidas.any():

        # Localiza os índices das datas inválidas
        indices_invalidos = dataframe_tratado.index[
            datas_limite_invalidas
        ].tolist()

        # Lança uma exceção com os índices problemáticos
        raise ValueError(
            f"A coluna '{coluna_data_limite}' possui datas inválidas "
            f"nos índices: {indices_invalidos}."
        )

    # Interrompe a execução quando existem conclusões inválidas
    if datas_conclusao_invalidas.any():

        # Localiza os índices das datas inválidas
        indices_invalidos = dataframe_tratado.index[
            datas_conclusao_invalidas
        ].tolist()

        # Lança uma exceção com os índices problemáticos
        raise ValueError(
            f"A coluna '{coluna_data_conclusao}' possui datas inválidas "
            f"nos índices: {indices_invalidos}."
        )

    # Compara a data de conclusão com a data limite
    registros_atrasados = data_conclusao > data_limite

    # Cria a coluna ATRASADA com SIM ou NÃO
    dataframe_tratado["ATRASADA"] = registros_atrasados.map(
        {
            True: "SIM",
            False: "NÃO"
        }
    )

    # Identifica linhas que possuem alguma das datas vazias
    datas_vazias = (
        data_limite.isna()
        | data_conclusao.isna()
    )

    # Mantém a situação vazia quando alguma data estiver ausente
    dataframe_tratado.loc[
        datas_vazias,
        "ATRASADA"
    ] = pd.NA

    # Retorna a cópia do DataFrame com a coluna ATRASADA
    return dataframe_tratado