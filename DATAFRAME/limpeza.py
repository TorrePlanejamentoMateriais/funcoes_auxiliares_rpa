"""
============================================================
Módulo: DATAFRAME / limpeza.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 24/09/2026
Última Alteração: 25/09/2026
Versão: 1.1.0

Descrição:
    Biblioteca de funções auxiliares para limpeza,
    validação e padronização de DataFrames utilizando pandas.

Funções disponíveis:
    - remover_espacos_colunas()
    - padronizar_nomes_colunas()
    - remover_duplicados()
    - remover_linhas_vazias()

Dependências:
    - pandas

Histórico:
    v1.0.0 - 24/09/2026
        - Criação inicial do módulo.
        - Implementação das funções de limpeza.

    v1.1.0 - 25/09/2026
        - Remoção do parâmetro booleano por_coluna.
        - Ajuste dos docstrings.
        - Inclusão de validações.
        - Correção da função remover_espacos_colunas.
        - Padronização dos nomes dos parâmetros.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa expressões regulares para tratamento de textos
import re

# Importa recursos para normalização e remoção de acentos
import unicodedata

# Importa a biblioteca pandas utilizando o apelido pd
import pandas as pd

from typing import Literal

# ----------------------------------------------------------------------------
# FUNÇÕES
# ----------------------------------------------------------------------------


def remover_espacos_colunas(
    dataframe: pd.DataFrame
) -> pd.DataFrame:
    """
    Remove espaços extras dos nomes das colunas de um DataFrame.

    A função realiza os seguintes tratamentos:

    - Converte os nomes das colunas para texto.
    - Remove espaços encontrados no início e no final.
    - Substitui múltiplos espaços internos por um único espaço.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá os espaços dos nomes das colunas removidos.

    Returns:
        pd.DataFrame:
            DataFrame com os nomes das colunas tratados.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        ValueError:
            Caso o DataFrame não possua nenhuma coluna.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame possui colunas
    if dataframe.columns.empty:
        raise ValueError(
            "O DataFrame não possui colunas."
        )

    # Cria uma cópia para não modificar diretamente o DataFrame original
    dataframe_tratado = dataframe.copy()

    # Converte os nomes para texto, remove espaços externos e espaços duplicados
    dataframe_tratado.columns = [
        re.sub(r"\s+", " ", str(coluna)).strip()
        for coluna in dataframe_tratado.columns
    ]

    # Retorna o DataFrame com os nomes das colunas tratados
    return dataframe_tratado


def padronizar_nomes_colunas(
    dataframe: pd.DataFrame
) -> pd.DataFrame:
    """
    Padroniza os nomes das colunas de um DataFrame.

    A função realiza os seguintes tratamentos:

    - Converte os nomes das colunas para texto.
    - Remove espaços encontrados no início e no final.
    - Converte todas as letras para minúsculas.
    - Remove acentos.
    - Substitui espaços e hífens por underline.
    - Remove caracteres especiais.
    - Remove underlines duplicados.
    - Remove underlines encontrados no início e no final.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá os nomes das colunas padronizados.

    Returns:
        pd.DataFrame:
            DataFrame com os nomes das colunas padronizados.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        ValueError:
            Caso o DataFrame não possua colunas ou a padronização
            gere nomes de colunas vazios ou duplicados.
    """

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se o DataFrame possui colunas
    if dataframe.columns.empty:
        raise ValueError(
            "O DataFrame não possui colunas."
        )

    # Cria uma cópia para não modificar diretamente o DataFrame original
    dataframe_tratado = dataframe.copy()

    # Cria uma lista para armazenar os nomes padronizados
    colunas_padronizadas = []

    # Percorre os nomes das colunas existentes no DataFrame
    for coluna in dataframe_tratado.columns:

        # Converte o nome para texto, remove espaços externos e usa minúsculas
        coluna_padronizada = str(coluna).strip().lower()

        # Separa os acentos das letras
        coluna_padronizada = unicodedata.normalize(
            "NFKD",
            coluna_padronizada
        )

        # Remove os caracteres responsáveis pela formação dos acentos
        coluna_padronizada = "".join(
            caractere
            for caractere in coluna_padronizada
            if not unicodedata.combining(caractere)
        )

        # Substitui grupos de espaços ou hífens por underline
        coluna_padronizada = re.sub(
            r"[\s-]+",
            "_",
            coluna_padronizada
        )

        # Remove caracteres que não sejam letras, números ou underline
        coluna_padronizada = re.sub(
            r"[^a-z0-9_]",
            "",
            coluna_padronizada
        )

        # Substitui underlines consecutivos por apenas um underline
        coluna_padronizada = re.sub(
            r"_+",
            "_",
            coluna_padronizada
        )

        # Remove underlines encontrados no início e no final
        coluna_padronizada = coluna_padronizada.strip("_")

        # Verifica se o nome da coluna ficou vazio após a padronização
        if not coluna_padronizada:
            raise ValueError(
                f"O nome da coluna '{coluna}' ficou vazio "
                "após a padronização."
            )

        # Adiciona o nome tratado à lista de colunas padronizadas
        colunas_padronizadas.append(coluna_padronizada)

    # Localiza nomes de colunas duplicados após a padronização
    colunas_duplicadas = {
        coluna
        for coluna in colunas_padronizadas
        if colunas_padronizadas.count(coluna) > 1
    }

    # Impede que o DataFrame fique com nomes de colunas duplicados
    if colunas_duplicadas:
        raise ValueError(
            "A padronização gerou nomes de colunas duplicados: "
            f"{sorted(colunas_duplicadas)}"
        )

    # Substitui os nomes atuais pelos nomes padronizados
    dataframe_tratado.columns = colunas_padronizadas

    # Retorna o DataFrame com os nomes das colunas padronizados
    return dataframe_tratado


def remover_duplicados(
    dataframe: pd.DataFrame,
    colunas: list[str] | None = None
) -> pd.DataFrame:
    """
    Remove registros duplicados de um DataFrame.

    Quando nenhuma coluna é informada, a função considera todas as
    colunas do DataFrame para identificar registros duplicados.

    Quando uma lista de colunas é informada, a função considera somente
    a combinação dos valores dessas colunas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá os registros duplicados removidos.

        colunas (list[str] | None, optional):
            Lista de colunas utilizadas para identificar os registros
            duplicados. O valor padrão é None.

    Returns:
        pd.DataFrame:
            DataFrame sem os registros duplicados.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame ou o parâmetro
            colunas não seja uma lista de strings.

        ValueError:
            Caso o DataFrame esteja vazio ou uma lista vazia seja informada.

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

    # Verifica se o argumento colunas é uma lista ou None
    if colunas is not None and not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings ou None."
        )

    # Verifica se foi informada uma lista vazia
    if colunas == []:
        raise ValueError(
            "A lista de colunas não pode estar vazia."
        )

    # Verifica se todos os itens da lista são strings
    if colunas is not None and not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Verifica quais colunas informadas não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas or []
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução caso existam colunas não encontradas
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Armazena a quantidade de linhas antes da remoção
    quantidade_antes = len(dataframe)

    # Cria uma cópia para não modificar diretamente o DataFrame original
    dataframe_tratado = dataframe.copy()

    # Remove registros duplicados com base nas colunas informadas
    if colunas is not None:
        dataframe_tratado = dataframe_tratado.drop_duplicates(
            subset=colunas
        )

    # Remove registros considerando todas as colunas do DataFrame
    else:
        dataframe_tratado = dataframe_tratado.drop_duplicates()

    # Reorganiza o índice depois da remoção das linhas
    dataframe_tratado = dataframe_tratado.reset_index(drop=True)

    # Armazena a quantidade de linhas depois da remoção
    quantidade_depois = len(dataframe_tratado)

    # Calcula a quantidade de registros removidos
    quantidade_removida = quantidade_antes - quantidade_depois

    # Exibe o resumo da operação realizada
    print(f"Quantidade de linhas antes: {quantidade_antes}")
    print(f"Quantidade de linhas depois: {quantidade_depois}")
    print(f"Quantidade removida: {quantidade_removida}")

    # Retorna o DataFrame sem os registros duplicados
    return dataframe_tratado


def remover_linhas_vazias(
    dataframe: pd.DataFrame,
    colunas: list[str] | None = None
) -> pd.DataFrame:
    """
    Remove linhas vazias ou com valores nulos de um DataFrame.

    Quando nenhuma coluna é informada, a função remove somente as linhas
    em que todas as colunas possuem valores nulos.

    Quando uma lista de colunas é informada, a função remove as linhas
    que possuam valor nulo em pelo menos uma das colunas especificadas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá as linhas vazias removidas.

        colunas (list[str] | None, optional):
            Lista de colunas que serão verificadas. O valor padrão é None.

    Returns:
        pd.DataFrame:
            DataFrame sem as linhas consideradas vazias.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame ou o parâmetro
            colunas não seja uma lista de strings.

        ValueError:
            Caso o DataFrame esteja vazio ou uma lista vazia seja informada.

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

    # Verifica se o argumento colunas é uma lista ou None
    if colunas is not None and not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings ou None."
        )

    # Verifica se foi informada uma lista vazia
    if colunas == []:
        raise ValueError(
            "A lista de colunas não pode estar vazia."
        )

    # Verifica se todos os itens da lista são strings
    if colunas is not None and not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Verifica quais colunas informadas não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas or []
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução caso existam colunas não encontradas
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Armazena a quantidade de linhas antes da remoção
    quantidade_antes = len(dataframe)

    # Cria uma cópia para não modificar diretamente o DataFrame original
    dataframe_tratado = dataframe.copy()

    # Remove linhas nulas com base nas colunas informadas
    if colunas is not None:
        dataframe_tratado = dataframe_tratado.dropna(
            subset=colunas
        )

    # Remove somente as linhas completamente nulas
    else:
        dataframe_tratado = dataframe_tratado.dropna(
            how="all"
        )

    # Reorganiza o índice depois da remoção das linhas
    dataframe_tratado = dataframe_tratado.reset_index(drop=True)

    # Armazena a quantidade de linhas depois da remoção
    quantidade_depois = len(dataframe_tratado)

    # Calcula a quantidade de linhas removidas
    quantidade_removida = quantidade_antes - quantidade_depois

    # Exibe o resumo da operação realizada
    print(f"Quantidade de linhas antes: {quantidade_antes}")
    print(f"Quantidade de linhas depois: {quantidade_depois}")
    print(f"Quantidade removida: {quantidade_removida}")

    # Retorna o DataFrame sem as linhas vazias
    return dataframe_tratado


def preencher_nulos(
    dataframe: pd.DataFrame,
    colunas: list[str] | None = None,
    metodo_preenchimento: Literal[
        "mediana",
        "interpolacao",
        "moda",
        "ffill"
    ] = "mediana"
) -> pd.DataFrame:
    """
    Preenche os valores nulos das colunas de um DataFrame.

    Métodos disponíveis:

    - mediana:
        Preenche os valores nulos com a mediana da coluna.
        Esse método pode ser utilizado somente em colunas numéricas.

    - interpolacao:
        Preenche os valores nulos utilizando interpolação linear.
        Esse método pode ser utilizado somente em colunas numéricas.

    - moda:
        Preenche os valores nulos com o valor mais frequente da coluna.
        Pode ser utilizado em colunas numéricas, textuais ou categóricas.

    - ffill:
        Preenche os valores nulos utilizando o último valor válido
        encontrado anteriormente na coluna.

    Quando nenhuma coluna for informada, todas as colunas do DataFrame
    serão processadas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que terá os valores nulos preenchidos.

        colunas (list[str] | None, optional):
            Lista das colunas que serão processadas.
            Quando o valor for None, todas as colunas do DataFrame
            serão processadas. O valor padrão é None.

        metodo_preenchimento (
            Literal["mediana", "interpolacao", "moda", "ffill"],
            optional
        ):
            Método utilizado para preencher os valores nulos.
            O valor padrão é "mediana".

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com os valores nulos preenchidos.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro colunas não seja uma lista ou None.

        TypeError:
            Caso algum item da lista de colunas não seja uma string.

        TypeError:
            Caso o método mediana ou interpolacao seja utilizado
            em uma coluna não numérica.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso seja informada uma lista de colunas vazia.

        ValueError:
            Caso o método de preenchimento seja inválido.

        ValueError:
            Caso não seja possível calcular a mediana ou a moda.

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

    # Verifica se o parâmetro colunas é uma lista ou None
    if colunas is not None and not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings ou None."
        )

    # Verifica se foi informada uma lista de colunas vazia
    if colunas == []:
        raise ValueError(
            "A lista de colunas não pode estar vazia."
        )

    # Verifica se todos os elementos da lista de colunas são strings
    if colunas is not None and not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Define todas as colunas quando nenhuma coluna específica for informada
    if colunas is None:
        colunas_selecionadas = list(dataframe.columns)
    else:
        colunas_selecionadas = colunas.copy()

    # Identifica as colunas informadas que não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas_selecionadas
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução caso alguma coluna não exista no DataFrame
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Verifica se o método de preenchimento informado é uma string
    if not isinstance(metodo_preenchimento, str):
        raise TypeError(
            "O parâmetro 'metodo_preenchimento' deve ser uma string."
        )

    # Padroniza o nome do método removendo espaços e convertendo para minúsculas
    metodo_preenchimento = metodo_preenchimento.strip().lower()

    # Define os métodos de preenchimento aceitos pela função
    metodos_permitidos = {
        "mediana",
        "interpolacao",
        "moda",
        "ffill"
    }

    # Verifica se o método de preenchimento informado é válido
    if metodo_preenchimento not in metodos_permitidos:
        raise ValueError(
            "Método de preenchimento inválido. "
            "Escolha uma das seguintes opções: "
            f"{sorted(metodos_permitidos)}."
        )

    # Cria uma cópia para preservar o DataFrame original
    dataframe_tratado = dataframe.copy()

    # Calcula a quantidade total de valores nulos antes do preenchimento
    quantidade_nulos_antes = int(
        dataframe_tratado[colunas_selecionadas]
        .isna()
        .sum()
        .sum()
    )

    # Cria um dicionário para armazenar a quantidade preenchida por coluna
    resumo_preenchimento = {}

    # Percorre todas as colunas selecionadas
    for coluna in colunas_selecionadas:

        # Calcula a quantidade de valores nulos da coluna antes do preenchimento
        quantidade_coluna_antes = int(
            dataframe_tratado[coluna].isna().sum()
        )

        # Registra zero e pula a coluna quando não existem valores nulos
        if quantidade_coluna_antes == 0:
            resumo_preenchimento[coluna] = 0
            continue

        # Executa o preenchimento utilizando a mediana
        if metodo_preenchimento == "mediana":

            # Verifica se a coluna possui valores numéricos
            if not pd.api.types.is_numeric_dtype(
                dataframe_tratado[coluna]
            ):
                raise TypeError(
                    f"A coluna '{coluna}' não é numérica e não pode "
                    "ser preenchida utilizando a mediana."
                )

            # Calcula a mediana dos valores válidos da coluna
            valor_mediana = dataframe_tratado[coluna].median()

            # Verifica se foi possível calcular a mediana
            if pd.isna(valor_mediana):
                raise ValueError(
                    f"Não foi possível calcular a mediana da coluna "
                    f"'{coluna}', pois ela não possui valores válidos."
                )

            # Preenche os valores nulos utilizando a mediana calculada
            dataframe_tratado[coluna] = (
                dataframe_tratado[coluna].fillna(valor_mediana)
            )

        # Executa o preenchimento utilizando interpolação linear
        elif metodo_preenchimento == "interpolacao":

            # Verifica se a coluna possui valores numéricos
            if not pd.api.types.is_numeric_dtype(
                dataframe_tratado[coluna]
            ):
                raise TypeError(
                    f"A coluna '{coluna}' não é numérica e não pode "
                    "ser preenchida utilizando interpolação."
                )

            # Verifica se a coluna possui pelo menos um valor válido
            if dataframe_tratado[coluna].notna().sum() == 0:
                raise ValueError(
                    f"Não foi possível aplicar interpolação na coluna "
                    f"'{coluna}', pois ela não possui valores válidos."
                )

            # Preenche os valores internos utilizando interpolação linear
            dataframe_tratado[coluna] = (
                dataframe_tratado[coluna].interpolate(
                    method="linear",
                    limit_direction="both"
                )
            )

        # Executa o preenchimento utilizando a moda
        elif metodo_preenchimento == "moda":

            # Calcula os valores mais frequentes da coluna
            valores_moda = dataframe_tratado[coluna].mode(
                dropna=True
            )

            # Verifica se foi possível calcular a moda
            if valores_moda.empty:
                raise ValueError(
                    f"Não foi possível calcular a moda da coluna "
                    f"'{coluna}', pois ela não possui valores válidos."
                )

            # Seleciona o primeiro valor quando houver mais de uma moda
            valor_moda = valores_moda.iloc[0]

            # Preenche os valores nulos utilizando a moda calculada
            dataframe_tratado[coluna] = (
                dataframe_tratado[coluna].fillna(valor_moda)
            )

        # Executa o preenchimento utilizando o último valor válido
        elif metodo_preenchimento == "ffill":

            # Preenche os nulos com o último valor válido anterior
            dataframe_tratado[coluna] = (
                dataframe_tratado[coluna].ffill()
            )

        # Calcula a quantidade de valores nulos restante na coluna
        quantidade_coluna_depois = int(
            dataframe_tratado[coluna].isna().sum()
        )

        # Calcula quantos valores foram preenchidos na coluna
        quantidade_preenchida_coluna = (
            quantidade_coluna_antes
            - quantidade_coluna_depois
        )

        # Armazena o resultado do preenchimento da coluna
        resumo_preenchimento[coluna] = quantidade_preenchida_coluna

    # Calcula a quantidade total de nulos depois do preenchimento
    quantidade_nulos_depois = int(
        dataframe_tratado[colunas_selecionadas]
        .isna()
        .sum()
        .sum()
    )

    # Calcula a quantidade total de valores preenchidos
    quantidade_total_preenchida = (
        quantidade_nulos_antes
        - quantidade_nulos_depois
    )

    # Exibe o método utilizado no preenchimento
    print(
        f"Método de preenchimento utilizado: "
        f"{metodo_preenchimento}"
    )

    # Exibe a quantidade de valores nulos antes do preenchimento
    print(
        f"Quantidade de valores nulos antes: "
        f"{quantidade_nulos_antes}"
    )

    # Exibe a quantidade de valores nulos depois do preenchimento
    print(
        f"Quantidade de valores nulos depois: "
        f"{quantidade_nulos_depois}"
    )

    # Exibe a quantidade total de valores preenchidos
    print(
        f"Quantidade total de valores preenchidos: "
        f"{quantidade_total_preenchida}"
    )

    # Exibe a quantidade de valores preenchidos em cada coluna
    for coluna, quantidade in resumo_preenchimento.items():
        print(
            f"Coluna '{coluna}': "
            f"{quantidade} valor(es) preenchido(s)."
        )

    # Retorna a cópia do DataFrame com os valores nulos preenchidos
    return dataframe_tratado

def converter_datas(
    dataframe: pd.DataFrame,
    colunas: list[str] | None = None,
    formato: Literal[
        "dia-mes-ano",
        "dia/mes/ano",
        "dia.mes.ano",
        "ano-mes-dia",
        "ano/mes/dia",
        "ano.mes.dia",
        "mes-dia-ano",
        "mes/dia/ano",
        "mes.dia.ano",
        "mes-ano",
        "mes/ano",
        "mes.ano",
        "ano-mes",
        "ano/mes",
        "ano.mes",
        "dia-mes",
        "dia/mes",
        "dia.mes",
        "mes-dia",
        "mes/dia",
        "mes.dia",
        "dia-mes-ano-hora",
        "dia/mes/ano hora",
        "ano-mes-dia-hora"
    ] = "dia-mes-ano",
    formato_entrada: str | None = None
) -> pd.DataFrame:
    """
    Converte e formata colunas de datas de um DataFrame.

    A função cria uma cópia do DataFrame original e converte as colunas
    selecionadas para o formato de data informado.

    Formatos disponíveis:

    - dia-mes-ano: 25-09-2026
    - dia/mes/ano: 25/09/2026
    - dia.mes.ano: 25.09.2026

    - ano-mes-dia: 2026-09-25
    - ano/mes/dia: 2026/09/25
    - ano.mes.dia: 2026.09.25

    - mes-dia-ano: 09-25-2026
    - mes/dia/ano: 09/25/2026
    - mes.dia.ano: 09.25.2026

    - mes-ano: 09-2026
    - mes/ano: 09/2026
    - mes.ano: 09.2026

    - ano-mes: 2026-09
    - ano/mes: 2026/09
    - ano.mes: 2026.09

    - dia-mes: 25-09
    - dia/mes: 25/09
    - dia.mes: 25.09

    - mes-dia: 09-25
    - mes/dia: 09/25
    - mes.dia: 09.25

    - dia-mes-ano-hora: 25-09-2026 14:30:00
    - dia/mes/ano hora: 25/09/2026 14:30:00
    - ano-mes-dia-hora: 2026-09-25 14:30:00

    Quando colunas for None, todas as colunas do DataFrame serão
    selecionadas para conversão.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que contém as colunas de data.

        colunas (list[str] | None, optional):
            Lista das colunas que serão convertidas.
            Quando for None, todas as colunas serão selecionadas.
            O valor padrão é None.

        formato (str, optional):
            Formato de saída que será aplicado às datas.
            O valor padrão é "dia-mes-ano".

        formato_entrada (str | None, optional):
            Formato original das datas, utilizando os códigos do Python.

            Exemplos:

            - "%d/%m/%Y" para 25/09/2026
            - "%Y-%m-%d" para 2026-09-25
            - "%d.%m.%Y" para 25.09.2026
            - "%d/%m/%Y %H:%M:%S" para data e hora

            Quando for None, o pandas tentará identificar o formato
            automaticamente. O valor padrão é None.

    Returns:
        pd.DataFrame:
            Cópia do DataFrame com as colunas de datas convertidas
            e formatadas.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o parâmetro colunas não seja uma lista ou None.

        TypeError:
            Caso algum elemento da lista de colunas não seja uma string.

        TypeError:
            Caso o formato informado não seja uma string.

        TypeError:
            Caso formato_entrada não seja uma string ou None.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso seja informada uma lista de colunas vazia.

        ValueError:
            Caso o formato de saída informado seja inválido.

        ValueError:
            Caso algum valor não possa ser convertido para data.

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

    # Verifica se o parâmetro colunas é uma lista ou None
    if colunas is not None and not isinstance(colunas, list):
        raise TypeError(
            "O parâmetro 'colunas' deve ser uma lista de strings ou None."
        )

    # Verifica se foi informada uma lista de colunas vazia
    if colunas == []:
        raise ValueError(
            "A lista de colunas não pode estar vazia."
        )

    # Verifica se todos os elementos da lista de colunas são strings
    if colunas is not None and not all(
        isinstance(coluna, str)
        for coluna in colunas
    ):
        raise TypeError(
            "Todos os itens da lista 'colunas' devem ser strings."
        )

    # Verifica se o formato informado é uma string
    if not isinstance(formato, str):
        raise TypeError(
            "O parâmetro 'formato' deve ser uma string."
        )

    # Verifica se o formato de entrada é uma string ou None
    if (
        formato_entrada is not None
        and not isinstance(formato_entrada, str)
    ):
        raise TypeError(
            "O parâmetro 'formato_entrada' deve ser uma string ou None."
        )

    # Define todas as colunas quando nenhuma coluna específica for informada
    if colunas is None:
        colunas_selecionadas = list(dataframe.columns)
    else:
        colunas_selecionadas = colunas.copy()

    # Identifica as colunas informadas que não existem no DataFrame
    colunas_nao_encontradas = [
        coluna
        for coluna in colunas_selecionadas
        if coluna not in dataframe.columns
    ]

    # Interrompe a execução quando alguma coluna não existe
    if colunas_nao_encontradas:
        raise KeyError(
            "Coluna(s) não encontrada(s) no DataFrame: "
            f"{colunas_nao_encontradas}"
        )

    # Padroniza o nome do formato informado
    formato = formato.strip().lower()

    # Relaciona os formatos amigáveis aos formatos utilizados pelo Python
    formatos_disponiveis = {
        "dia-mes-ano": "%d-%m-%Y",
        "dia/mes/ano": "%d/%m/%Y",
        "dia.mes.ano": "%d.%m.%Y",

        "ano-mes-dia": "%Y-%m-%d",
        "ano/mes/dia": "%Y/%m/%d",
        "ano.mes.dia": "%Y.%m.%d",

        "mes-dia-ano": "%m-%d-%Y",
        "mes/dia/ano": "%m/%d/%Y",
        "mes.dia.ano": "%m.%d.%Y",

        "mes-ano": "%m-%Y",
        "mes/ano": "%m/%Y",
        "mes.ano": "%m.%Y",

        "ano-mes": "%Y-%m",
        "ano/mes": "%Y/%m",
        "ano.mes": "%Y.%m",

        "dia-mes": "%d-%m",
        "dia/mes": "%d/%m",
        "dia.mes": "%d.%m",

        "mes-dia": "%m-%d",
        "mes/dia": "%m/%d",
        "mes.dia": "%m.%d",

        "dia-mes-ano-hora": "%d-%m-%Y %H:%M:%S",
        "dia/mes/ano hora": "%d/%m/%Y %H:%M:%S",
        "ano-mes-dia-hora": "%Y-%m-%d %H:%M:%S",
    }

    # Verifica se o formato de saída informado está disponível
    if formato not in formatos_disponiveis:
        raise ValueError(
            "Formato de data inválido. "
            "Escolha uma das seguintes opções: "
            f"{sorted(formatos_disponiveis.keys())}."
        )

    # Obtém o código de formatação correspondente ao formato escolhido
    codigo_formato_saida = formatos_disponiveis[formato]

    # Cria uma cópia profunda para preservar o DataFrame original
    dataframe_tratado = dataframe.copy(deep=True)

    # Cria um dicionário para armazenar o resumo da conversão
    resumo_conversao = {}

    # Percorre todas as colunas selecionadas
    for coluna in colunas_selecionadas:

        # Conta a quantidade de valores não nulos antes da conversão
        quantidade_valores = int(
            dataframe_tratado[coluna].notna().sum()
        )

        # Ignora a coluna quando ela não possui valores preenchidos
        if quantidade_valores == 0:
            resumo_conversao[coluna] = 0
            continue

        try:
            # Converte utilizando um formato de entrada específico
            if formato_entrada is not None:
                coluna_convertida = pd.to_datetime(
                    dataframe_tratado[coluna],
                    format=formato_entrada,
                    errors="raise"
                )

            # Converte tentando identificar automaticamente o formato
            else:
                coluna_convertida = pd.to_datetime(
                    dataframe_tratado[coluna],
                    errors="raise",
                    dayfirst=True
                )

        except (ValueError, TypeError) as erro:
            raise ValueError(
                f"Não foi possível converter a coluna '{coluna}' "
                f"para data. Verifique os valores e o formato de entrada. "
                f"Erro original: {erro}"
            ) from erro

        # Formata as datas convertidas no padrão de saída escolhido
        coluna_formatada = coluna_convertida.dt.strftime(
            codigo_formato_saida
        )

        # Mantém os valores originalmente nulos como valores ausentes
        coluna_formatada = coluna_formatada.where(
            dataframe_tratado[coluna].notna(),
            pd.NA
        )

        # Substitui a coluna original pela coluna formatada
        dataframe_tratado[coluna] = coluna_formatada

        # Registra a quantidade de datas convertidas na coluna
        resumo_conversao[coluna] = quantidade_valores

    # Exibe o formato aplicado às colunas selecionadas
    print(f"Formato de saída utilizado: {formato}")

    # Exibe os resultados de cada coluna processada
    for coluna, quantidade in resumo_conversao.items():
        print(
            f"Coluna '{coluna}': "
            f"{quantidade} valor(es) convertido(s)."
        )

    # Retorna a cópia do DataFrame com as datas formatadas
    return dataframe_tratado