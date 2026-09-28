"""
============================================================
Módulo: DATAFRAME / exportacao.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 25/09/2026
Última Alteração: 25/09/2026
Versão: 1.0.0

Descrição:
    Biblioteca de funções auxiliares para exportação de
    DataFrames nos formatos Excel, CSV e Parquet
    utilizando pandas.

Funções disponíveis:
    - exportar_excel()
    - exportar_csv()
    - exportar_parquet()

Dependências:
    - pandas
    - openpyxl
    - pyarrow

Histórico:
    v1.0.0 - 25/09/2026
        - Criação inicial do módulo.
        - Implementação da função exportar_excel.
        - Implementação da função exportar_csv.
        - Implementação da função exportar_parquet.
        - Inclusão de validações do DataFrame.
        - Inclusão de validações dos caminhos de destino.
        - Inclusão da criação automática das pastas.
        - Inclusão de tratamentos específicos de exceções.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa o módulo os para manipulação de caminhos e diretórios
import os

# Importa a biblioteca pandas utilizando o apelido pd
import pandas as pd


# ----------------------------------------------------------------------------
# FUNÇÕES
# ----------------------------------------------------------------------------


def exportar_excel(
    dataframe: pd.DataFrame,
    caminho: str
) -> None:
    """
    Exporta um DataFrame para um arquivo Excel.

    A função valida o DataFrame e o caminho informado, cria a pasta
    de destino quando ela não existir e exporta o arquivo sem incluir
    o índice do DataFrame.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.

        caminho (str):
            Caminho completo do arquivo Excel de destino.

            Exemplo:
                C:\\arquivos\\relatorio.xlsx

    Returns:
        None:
            A função não retorna valores.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o caminho informado não seja uma string.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso o caminho esteja vazio.

        ValueError:
            Caso o caminho não possua a extensão .xlsx.

        PermissionError:
            Caso o arquivo esteja aberto, bloqueado ou o usuário
            não possua permissão de escrita.

        ModuleNotFoundError:
            Caso a biblioteca openpyxl não esteja instalada.

        OSError:
            Caso não seja possível criar a pasta de destino ou
            salvar o arquivo.

        RuntimeError:
            Caso ocorra outro erro durante a exportação.
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

    # Verifica se o caminho informado é uma string
    if not isinstance(caminho, str):
        raise TypeError(
            "O parâmetro 'caminho' deve ser uma string."
        )

    # Verifica se o caminho informado está vazio
    if not caminho.strip():
        raise ValueError(
            "O parâmetro 'caminho' não pode estar vazio."
        )

    # Remove espaços externos, expande o diretório do usuário
    # e transforma o caminho em um caminho absoluto
    caminho_normalizado = os.path.abspath(
        os.path.expanduser(
            caminho.strip()
        )
    )

    # Verifica se o caminho possui a extensão .xlsx
    if not caminho_normalizado.lower().endswith(".xlsx"):
        raise ValueError(
            "O caminho de destino deve possuir a extensão '.xlsx'."
        )

    # Obtém o caminho completo da pasta de destino
    pasta_destino = os.path.dirname(
        caminho_normalizado
    )

    try:

        # Cria a pasta de destino quando ela ainda não existir
        if pasta_destino:
            os.makedirs(
                pasta_destino,
                exist_ok=True
            )

    except OSError as erro:
        raise OSError(
            f"Não foi possível criar a pasta de destino "
            f"'{pasta_destino}': {erro}"
        ) from erro

    try:

        # Exporta o DataFrame para um arquivo Excel
        dataframe.to_excel(
            caminho_normalizado,
            index=False,
            engine="openpyxl"
        )

    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível salvar o arquivo "
            f"'{caminho_normalizado}'. Verifique se o arquivo está "
            "aberto no Excel ou se você possui permissão de escrita."
        ) from erro

    except ImportError as erro:
        raise ModuleNotFoundError(
            "A biblioteca 'openpyxl' não está instalada. "
            "Execute o comando: python -m pip install openpyxl"
        ) from erro

    except OSError as erro:
        raise OSError(
            f"Erro de sistema ao exportar o arquivo Excel "
            f"'{caminho_normalizado}': {erro}"
        ) from erro

    except Exception as erro:
        raise RuntimeError(
            "Erro inesperado ao exportar o DataFrame para Excel: "
            f"{erro}"
        ) from erro

    # Exibe uma mensagem após a exportação ser concluída
    print(
        "Arquivo Excel exportado com sucesso: "
        f"{caminho_normalizado}"
    )


def exportar_csv(
    dataframe: pd.DataFrame,
    caminho: str,
    separador: str = ";",
    codificacao: str = "utf-8-sig"
) -> None:
    """
    Exporta um DataFrame para um arquivo CSV.

    A função valida o DataFrame e o caminho informado, cria a pasta
    de destino quando ela não existir e exporta o arquivo sem incluir
    o índice do DataFrame.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.

        caminho (str):
            Caminho completo do arquivo CSV de destino.

            Exemplo:
                C:\\arquivos\\relatorio.csv

        separador (str, optional):
            Caractere utilizado para separar as colunas no arquivo.
            O valor padrão é ";".

        codificacao (str, optional):
            Codificação utilizada na criação do arquivo.
            O valor padrão é "utf-8-sig".

    Returns:
        None:
            A função não retorna valores.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso caminho, separador ou codificacao não sejam strings.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso caminho, separador ou codificacao estejam vazios.

        ValueError:
            Caso o separador possua mais de um caractere.

        ValueError:
            Caso o caminho não possua a extensão .csv.

        ValueError:
            Caso a codificação informada seja inválida.

        PermissionError:
            Caso o arquivo esteja aberto, bloqueado ou o usuário
            não possua permissão de escrita.

        OSError:
            Caso não seja possível criar a pasta de destino ou
            salvar o arquivo.

        RuntimeError:
            Caso ocorra outro erro durante a exportação.
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

    # Verifica se o caminho informado é uma string
    if not isinstance(caminho, str):
        raise TypeError(
            "O parâmetro 'caminho' deve ser uma string."
        )

    # Verifica se o caminho informado está vazio
    if not caminho.strip():
        raise ValueError(
            "O parâmetro 'caminho' não pode estar vazio."
        )

    # Verifica se o separador informado é uma string
    if not isinstance(separador, str):
        raise TypeError(
            "O parâmetro 'separador' deve ser uma string."
        )

    # Verifica se o separador informado está vazio
    if not separador:
        raise ValueError(
            "O parâmetro 'separador' não pode estar vazio."
        )

    # Verifica se o separador possui somente um caractere
    if len(separador) != 1:
        raise ValueError(
            "O parâmetro 'separador' deve possuir apenas um caractere."
        )

    # Verifica se a codificação informada é uma string
    if not isinstance(codificacao, str):
        raise TypeError(
            "O parâmetro 'codificacao' deve ser uma string."
        )

    # Verifica se a codificação informada está vazia
    if not codificacao.strip():
        raise ValueError(
            "O parâmetro 'codificacao' não pode estar vazio."
        )

    # Padroniza a codificação informada
    codificacao = codificacao.strip()

    # Remove espaços externos, expande o diretório do usuário
    # e transforma o caminho em um caminho absoluto
    caminho_normalizado = os.path.abspath(
        os.path.expanduser(
            caminho.strip()
        )
    )

    # Verifica se o caminho possui a extensão .csv
    if not caminho_normalizado.lower().endswith(".csv"):
        raise ValueError(
            "O caminho de destino deve possuir a extensão '.csv'."
        )

    # Obtém o caminho completo da pasta de destino
    pasta_destino = os.path.dirname(
        caminho_normalizado
    )

    try:

        # Cria a pasta de destino quando ela ainda não existir
        if pasta_destino:
            os.makedirs(
                pasta_destino,
                exist_ok=True
            )

    except OSError as erro:
        raise OSError(
            f"Não foi possível criar a pasta de destino "
            f"'{pasta_destino}': {erro}"
        ) from erro

    try:

        # Exporta o DataFrame para um arquivo CSV
        dataframe.to_csv(
            caminho_normalizado,
            index=False,
            sep=separador,
            encoding=codificacao
        )

    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível salvar o arquivo "
            f"'{caminho_normalizado}'. Verifique se o arquivo está "
            "aberto em outro programa ou se você possui permissão "
            "de escrita."
        ) from erro

    except LookupError as erro:
        raise ValueError(
            f"A codificação '{codificacao}' não é válida."
        ) from erro

    except OSError as erro:
        raise OSError(
            f"Erro de sistema ao exportar o arquivo CSV "
            f"'{caminho_normalizado}': {erro}"
        ) from erro

    except Exception as erro:
        raise RuntimeError(
            "Erro inesperado ao exportar o DataFrame para CSV: "
            f"{erro}"
        ) from erro

    # Exibe uma mensagem após a exportação ser concluída
    print(
        "Arquivo CSV exportado com sucesso: "
        f"{caminho_normalizado}"
    )


def exportar_parquet(
    dataframe: pd.DataFrame,
    caminho: str,
    compressao: str | None = "snappy"
) -> None:
    """
    Exporta um DataFrame para um arquivo Parquet.

    A função valida o DataFrame e o caminho informado, cria a pasta
    de destino quando ela não existir e exporta o arquivo sem incluir
    o índice do DataFrame.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.

        caminho (str):
            Caminho completo do arquivo Parquet de destino.

            Exemplo:
                C:\\arquivos\\relatorio.parquet

        compressao (str | None, optional):
            Método de compressão utilizado no arquivo.

            Valores disponíveis:

            - "snappy"
            - "gzip"
            - "brotli"
            - "lz4"
            - "zstd"
            - None

            O valor padrão é "snappy".

    Returns:
        None:
            A função não retorna valores.

    Raises:
        TypeError:
            Caso o objeto informado não seja um DataFrame.

        TypeError:
            Caso o caminho não seja uma string.

        TypeError:
            Caso compressao não seja uma string ou None.

        ValueError:
            Caso o DataFrame esteja vazio.

        ValueError:
            Caso o caminho esteja vazio.

        ValueError:
            Caso o caminho não possua a extensão .parquet.

        ValueError:
            Caso o método de compressão seja inválido.

        PermissionError:
            Caso o arquivo esteja bloqueado ou o usuário não possua
            permissão de escrita.

        ModuleNotFoundError:
            Caso a biblioteca pyarrow não esteja instalada.

        OSError:
            Caso não seja possível criar a pasta de destino ou
            salvar o arquivo.

        RuntimeError:
            Caso ocorra outro erro durante a exportação.
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

    # Verifica se o caminho informado é uma string
    if not isinstance(caminho, str):
        raise TypeError(
            "O parâmetro 'caminho' deve ser uma string."
        )

    # Verifica se o caminho informado está vazio
    if not caminho.strip():
        raise ValueError(
            "O parâmetro 'caminho' não pode estar vazio."
        )

    # Verifica se a compressão é uma string ou None
    if (
        compressao is not None
        and not isinstance(compressao, str)
    ):
        raise TypeError(
            "O parâmetro 'compressao' deve ser uma string ou None."
        )

    # Padroniza o nome da compressão quando ela for informada
    if compressao is not None:
        compressao = compressao.strip().lower()

    # Define os métodos de compressão aceitos
    compressoes_permitidas = {
        "snappy",
        "gzip",
        "brotli",
        "lz4",
        "zstd"
    }

    # Verifica se a compressão informada é válida
    if (
        compressao is not None
        and compressao not in compressoes_permitidas
    ):
        raise ValueError(
            "Método de compressão inválido. "
            "Escolha uma das seguintes opções: "
            f"{sorted(compressoes_permitidas)} ou None."
        )

    # Remove espaços externos, expande o diretório do usuário
    # e transforma o caminho em um caminho absoluto
    caminho_normalizado = os.path.abspath(
        os.path.expanduser(
            caminho.strip()
        )
    )

    # Verifica se o caminho possui a extensão .parquet
    if not caminho_normalizado.lower().endswith(".parquet"):
        raise ValueError(
            "O caminho de destino deve possuir a extensão '.parquet'."
        )

    # Obtém o caminho completo da pasta de destino
    pasta_destino = os.path.dirname(
        caminho_normalizado
    )

    try:

        # Cria a pasta de destino quando ela ainda não existir
        if pasta_destino:
            os.makedirs(
                pasta_destino,
                exist_ok=True
            )

    except OSError as erro:
        raise OSError(
            f"Não foi possível criar a pasta de destino "
            f"'{pasta_destino}': {erro}"
        ) from erro

    try:

        # Exporta o DataFrame para um arquivo Parquet
        dataframe.to_parquet(
            caminho_normalizado,
            index=False,
            engine="pyarrow",
            compression=compressao
        )

    except PermissionError as erro:
        raise PermissionError(
            f"Não foi possível salvar o arquivo "
            f"'{caminho_normalizado}'. Verifique se o arquivo está "
            "bloqueado ou se você possui permissão de escrita."
        ) from erro

    except ImportError as erro:
        raise ModuleNotFoundError(
            "A biblioteca 'pyarrow' não está instalada. "
            "Execute o comando: python -m pip install pyarrow"
        ) from erro

    except OSError as erro:
        raise OSError(
            f"Erro de sistema ao exportar o arquivo Parquet "
            f"'{caminho_normalizado}': {erro}"
        ) from erro

    except Exception as erro:
        raise RuntimeError(
            "Erro inesperado ao exportar o DataFrame para Parquet: "
            f"{erro}"
        ) from erro

    # Exibe uma mensagem após a exportação ser concluída
    print(
        "Arquivo Parquet exportado com sucesso: "
        f"{caminho_normalizado}"
    )