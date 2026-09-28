"""
============================================================
Modulo: test_exportacao.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 25/09/2026
Ultima Alteracao: 25/09/2026
Versao: 1.0.0

Descricao:
    Testes unitarios das funcoes presentes no modulo
    DATAFRAME/exportacao.py.

Funcoes testadas:
    - exportar_excel()
    - exportar_csv()
    - exportar_parquet()

Dependencias:
    - pandas
    - pytest

Historico:
    v1.0.0 - 25/09/2026
        - Criacao inicial dos testes unitarios.
        - Testes de validacao dos parametros.
        - Testes de criacao das pastas de destino.
        - Testes das chamadas de exportacao.
        - Testes de tratamento das excecoes.
        - Testes de preservacao do DataFrame original.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

from unittest.mock import patch

import pandas as pd
import pytest

from exportacao import (
    exportar_csv,
    exportar_excel,
    exportar_parquet,
)


# ----------------------------------------------------------------------------
# FIXTURES
# ----------------------------------------------------------------------------


@pytest.fixture
def dataframe_valido():
    """Cria um DataFrame valido para os testes."""

    return pd.DataFrame(
        {
            "pedido": [100, 200],
            "fornecedor": ["Fornecedor A", "Fornecedor B"],
            "valor": [150.50, 300.75],
        }
    )


@pytest.fixture
def dataframe_vazio():
    """Cria um DataFrame vazio para os testes."""

    return pd.DataFrame(columns=["pedido", "fornecedor", "valor"])


# ----------------------------------------------------------------------------
# TESTES DA FUNCAO exportar_excel()
# ----------------------------------------------------------------------------


def test_exportar_excel_chama_to_excel_corretamente(
    dataframe_valido,
    tmp_path,
):
    """Verifica os argumentos utilizados para exportar Excel."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(pd.DataFrame, "to_excel") as mock_to_excel:
        resultado = exportar_excel(dataframe_valido, str(caminho))

    mock_to_excel.assert_called_once_with(
        str(caminho.resolve()),
        index=False,
        engine="openpyxl",
    )
    assert resultado is None


def test_exportar_excel_cria_pasta_destino(dataframe_valido, tmp_path):
    """Verifica se a pasta de destino e criada automaticamente."""

    caminho = tmp_path / "nova_pasta" / "relatorio.xlsx"

    with patch.object(pd.DataFrame, "to_excel"):
        exportar_excel(dataframe_valido, str(caminho))

    assert caminho.parent.exists()
    assert caminho.parent.is_dir()


def test_exportar_excel_exibe_mensagem_sucesso(
    dataframe_valido,
    tmp_path,
    capsys,
):
    """Verifica a mensagem exibida apos a exportacao."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(pd.DataFrame, "to_excel"):
        exportar_excel(dataframe_valido, str(caminho))

    saida = capsys.readouterr().out
    assert "Arquivo Excel exportado com sucesso" in saida
    assert str(caminho.resolve()) in saida


def test_exportar_excel_nao_altera_dataframe_original(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a exportacao nao altera o DataFrame original."""

    original = dataframe_valido.copy(deep=True)
    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(pd.DataFrame, "to_excel"):
        exportar_excel(dataframe_valido, str(caminho))

    pd.testing.assert_frame_equal(dataframe_valido, original)


@pytest.mark.parametrize(
    "objeto_invalido",
    [None, [], {}, "texto", 10, 10.5],
)
def test_exportar_excel_dataframe_invalido(objeto_invalido, tmp_path):
    """Verifica o erro quando o objeto não é um DataFrame."""

    caminho = tmp_path / "relatorio.xlsx"

    with pytest.raises(TypeError, match="não é um DataFrame"):
        exportar_excel(objeto_invalido, str(caminho))


def test_exportar_excel_dataframe_vazio(dataframe_vazio, tmp_path):
    """Verifica o erro quando o DataFrame está vazio."""

    caminho = tmp_path / "relatorio.xlsx"

    with pytest.raises(ValueError, match="DataFrame está vazio"):
        exportar_excel(dataframe_vazio, str(caminho))


@pytest.mark.parametrize(
    "caminho_invalido",
    [None, 10, 10.5, [], {}, ("arquivo.xlsx",)],
)
def test_exportar_excel_caminho_nao_e_string(
    dataframe_valido,
    caminho_invalido,
):
    """Verifica o erro quando o caminho nao e uma string."""

    with pytest.raises(TypeError, match="'caminho'.*string"):
        exportar_excel(dataframe_valido, caminho_invalido)


@pytest.mark.parametrize("caminho_vazio", ["", " ", "   ", "\t", "\n"])
def test_exportar_excel_caminho_vazio(dataframe_valido, caminho_vazio):
    """Verifica o erro quando o caminho esta vazio."""

    with pytest.raises(ValueError, match="'caminho'.*vazio"):
        exportar_excel(dataframe_valido, caminho_vazio)


@pytest.mark.parametrize(
    "extensao_invalida",
    ["relatorio.xls", "relatorio.csv", "relatorio.txt", "relatorio"],
)
def test_exportar_excel_extensao_invalida(
    dataframe_valido,
    tmp_path,
    extensao_invalida,
):
    """Verifica o erro quando a extensao nao e XLSX."""

    caminho = tmp_path / extensao_invalida

    with pytest.raises(ValueError, match="extensão '.xlsx'"):
        exportar_excel(dataframe_valido, str(caminho))


def test_exportar_excel_aceita_extensao_maiuscula(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a extensao XLSX em maiusculo e aceita."""

    caminho = tmp_path / "relatorio.XLSX"

    with patch.object(pd.DataFrame, "to_excel") as mock_to_excel:
        exportar_excel(dataframe_valido, str(caminho))

    assert mock_to_excel.call_count == 1


def test_exportar_excel_erro_ao_criar_pasta(dataframe_valido, tmp_path):
    """Verifica o tratamento de erro na criacao da pasta."""

    caminho = tmp_path / "pasta" / "relatorio.xlsx"

    with patch("exportacao.os.makedirs", side_effect=OSError("falha")):
        with pytest.raises(OSError, match="criar a pasta de destino"):
            exportar_excel(dataframe_valido, str(caminho))


def test_exportar_excel_erro_de_permissao(dataframe_valido, tmp_path):
    """Verifica o tratamento de PermissionError."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(
        pd.DataFrame,
        "to_excel",
        side_effect=PermissionError("arquivo bloqueado"),
    ):
        with pytest.raises(PermissionError, match="permissão de escrita"):
            exportar_excel(dataframe_valido, str(caminho))


def test_exportar_excel_openpyxl_ausente(dataframe_valido, tmp_path):
    """Verifica o tratamento da ausencia do openpyxl."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(
        pd.DataFrame,
        "to_excel",
        side_effect=ImportError("openpyxl ausente"),
    ):
        with pytest.raises(ModuleNotFoundError, match="openpyxl"):
            exportar_excel(dataframe_valido, str(caminho))


def test_exportar_excel_erro_de_sistema(dataframe_valido, tmp_path):
    """Verifica o tratamento de OSError durante a exportacao."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(
        pd.DataFrame,
        "to_excel",
        side_effect=OSError("falha no disco"),
    ):
        with pytest.raises(OSError, match="Erro de sistema"):
            exportar_excel(dataframe_valido, str(caminho))


def test_exportar_excel_erro_inesperado(dataframe_valido, tmp_path):
    """Verifica o tratamento de erros inesperados."""

    caminho = tmp_path / "relatorio.xlsx"

    with patch.object(
        pd.DataFrame,
        "to_excel",
        side_effect=RuntimeError("falha inesperada"),
    ):
        with pytest.raises(RuntimeError, match="Erro inesperado"):
            exportar_excel(dataframe_valido, str(caminho))


# ----------------------------------------------------------------------------
# TESTES DA FUNCAO exportar_csv()
# ----------------------------------------------------------------------------


def test_exportar_csv_cria_arquivo_real(dataframe_valido, tmp_path):
    """Verifica a criacao real de um arquivo CSV."""

    caminho = tmp_path / "relatorio.csv"

    resultado = exportar_csv(dataframe_valido, str(caminho))

    assert resultado is None
    assert caminho.exists()
    assert caminho.is_file()


def test_exportar_csv_conteudo_padrao(dataframe_valido, tmp_path):
    """Verifica separador, codificacao e ausencia do indice."""

    caminho = tmp_path / "relatorio.csv"

    exportar_csv(dataframe_valido, str(caminho))

    resultado = pd.read_csv(
        caminho,
        sep=";",
        encoding="utf-8-sig",
    )

    pd.testing.assert_frame_equal(resultado, dataframe_valido)
    assert "Unnamed: 0" not in resultado.columns


@pytest.mark.parametrize("separador", [",", ";", "|", "\t"])
def test_exportar_csv_separadores_validos(
    dataframe_valido,
    tmp_path,
    separador,
):
    """Verifica diferentes separadores de um caractere."""

    caminho = tmp_path / f"relatorio_{ord(separador)}.csv"

    exportar_csv(
        dataframe_valido,
        str(caminho),
        separador=separador,
    )

    resultado = pd.read_csv(
        caminho,
        sep=separador,
        encoding="utf-8-sig",
    )

    pd.testing.assert_frame_equal(resultado, dataframe_valido)


@pytest.mark.parametrize("codificacao", ["utf-8", "utf-8-sig", "latin-1"])
def test_exportar_csv_codificacoes_validas(
    dataframe_valido,
    tmp_path,
    codificacao,
):
    """Verifica diferentes codificacoes validas."""

    caminho = tmp_path / f"relatorio_{codificacao.replace('-', '_')}.csv"

    exportar_csv(
        dataframe_valido,
        str(caminho),
        codificacao=codificacao,
    )

    assert caminho.exists()


def test_exportar_csv_cria_pasta_destino(dataframe_valido, tmp_path):
    """Verifica se a pasta de destino e criada automaticamente."""

    caminho = tmp_path / "nova_pasta" / "relatorio.csv"

    exportar_csv(dataframe_valido, str(caminho))

    assert caminho.parent.exists()
    assert caminho.exists()


def test_exportar_csv_exibe_mensagem_sucesso(
    dataframe_valido,
    tmp_path,
    capsys,
):
    """Verifica a mensagem exibida apos a exportacao."""

    caminho = tmp_path / "relatorio.csv"

    exportar_csv(dataframe_valido, str(caminho))

    saida = capsys.readouterr().out
    assert "Arquivo CSV exportado com sucesso" in saida
    assert str(caminho.resolve()) in saida


def test_exportar_csv_nao_altera_dataframe_original(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a exportacao nao altera o DataFrame original."""

    original = dataframe_valido.copy(deep=True)
    caminho = tmp_path / "relatorio.csv"

    exportar_csv(dataframe_valido, str(caminho))

    pd.testing.assert_frame_equal(dataframe_valido, original)


@pytest.mark.parametrize(
    "objeto_invalido",
    [None, [], {}, "texto", 10, 10.5],
)
def test_exportar_csv_dataframe_invalido(objeto_invalido, tmp_path):
    """Verifica o erro quando o objeto não é um DataFrame."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(TypeError, match="não é um DataFrame"):
        exportar_csv(objeto_invalido, str(caminho))


def test_exportar_csv_dataframe_vazio(dataframe_vazio, tmp_path):
    """Verifica o erro quando o DataFrame está vazio."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(ValueError, match="DataFrame está vazio"):
        exportar_csv(dataframe_vazio, str(caminho))


@pytest.mark.parametrize(
    "caminho_invalido",
    [None, 10, 10.5, [], {}, ("arquivo.csv",)],
)
def test_exportar_csv_caminho_nao_e_string(
    dataframe_valido,
    caminho_invalido,
):
    """Verifica o erro quando o caminho nao e uma string."""

    with pytest.raises(TypeError, match="'caminho'.*string"):
        exportar_csv(dataframe_valido, caminho_invalido)


@pytest.mark.parametrize("caminho_vazio", ["", " ", "   ", "\t", "\n"])
def test_exportar_csv_caminho_vazio(dataframe_valido, caminho_vazio):
    """Verifica o erro quando o caminho esta vazio."""

    with pytest.raises(ValueError, match="'caminho'.*vazio"):
        exportar_csv(dataframe_valido, caminho_vazio)


@pytest.mark.parametrize(
    "extensao_invalida",
    ["relatorio.xlsx", "relatorio.txt", "relatorio.parquet", "relatorio"],
)
def test_exportar_csv_extensao_invalida(
    dataframe_valido,
    tmp_path,
    extensao_invalida,
):
    """Verifica o erro quando a extensao nao e CSV."""

    caminho = tmp_path / extensao_invalida

    with pytest.raises(ValueError, match="extensão '.csv'"):
        exportar_csv(dataframe_valido, str(caminho))


def test_exportar_csv_aceita_extensao_maiuscula(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a extensao CSV em maiusculo e aceita."""

    caminho = tmp_path / "relatorio.CSV"

    exportar_csv(dataframe_valido, str(caminho))

    assert caminho.exists()


@pytest.mark.parametrize("separador_invalido", [None, 10, [], {}, (",",)])
def test_exportar_csv_separador_nao_e_string(
    dataframe_valido,
    tmp_path,
    separador_invalido,
):
    """Verifica o erro quando o separador nao e uma string."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(TypeError, match="'separador'.*string"):
        exportar_csv(
            dataframe_valido,
            str(caminho),
            separador=separador_invalido,
        )


def test_exportar_csv_separador_vazio(dataframe_valido, tmp_path):
    """Verifica o erro quando o separador esta vazio."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(ValueError, match="'separador'.*vazio"):
        exportar_csv(dataframe_valido, str(caminho), separador="")


@pytest.mark.parametrize("separador_invalido", [";;", "sep", ", ", "\t\t"])
def test_exportar_csv_separador_com_mais_de_um_caractere(
    dataframe_valido,
    tmp_path,
    separador_invalido,
):
    """Verifica o erro quando o separador tem mais de um caractere."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(ValueError, match="apenas um caractere"):
        exportar_csv(
            dataframe_valido,
            str(caminho),
            separador=separador_invalido,
        )


@pytest.mark.parametrize("codificacao_invalida", [None, 10, [], {}, ("utf-8",)])
def test_exportar_csv_codificacao_nao_e_string(
    dataframe_valido,
    tmp_path,
    codificacao_invalida,
):
    """Verifica o erro quando a codificacao nao e uma string."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(TypeError, match="'codificacao'.*string"):
        exportar_csv(
            dataframe_valido,
            str(caminho),
            codificacao=codificacao_invalida,
        )


@pytest.mark.parametrize("codificacao_vazia", ["", " ", "   ", "\t", "\n"])
def test_exportar_csv_codificacao_vazia(
    dataframe_valido,
    tmp_path,
    codificacao_vazia,
):
    """Verifica o erro quando a codificacao esta vazia."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(ValueError, match="'codificacao'.*vazio"):
        exportar_csv(
            dataframe_valido,
            str(caminho),
            codificacao=codificacao_vazia,
        )


def test_exportar_csv_codificacao_inexistente(dataframe_valido, tmp_path):
    """Verifica o tratamento de uma codificacao inexistente."""

    caminho = tmp_path / "relatorio.csv"

    with pytest.raises(ValueError, match="codificação.*não é válida"):
        exportar_csv(
            dataframe_valido,
            str(caminho),
            codificacao="codificacao-inexistente",
        )


def test_exportar_csv_erro_ao_criar_pasta(dataframe_valido, tmp_path):
    """Verifica o tratamento de erro na criacao da pasta."""

    caminho = tmp_path / "pasta" / "relatorio.csv"

    with patch("exportacao.os.makedirs", side_effect=OSError("falha")):
        with pytest.raises(OSError, match="criar a pasta de destino"):
            exportar_csv(dataframe_valido, str(caminho))


def test_exportar_csv_erro_de_permissao(dataframe_valido, tmp_path):
    """Verifica o tratamento de PermissionError."""

    caminho = tmp_path / "relatorio.csv"

    with patch.object(
        pd.DataFrame,
        "to_csv",
        side_effect=PermissionError("arquivo bloqueado"),
    ):
        with pytest.raises(PermissionError, match="permissão de escrita"):
            exportar_csv(dataframe_valido, str(caminho))


def test_exportar_csv_erro_de_sistema(dataframe_valido, tmp_path):
    """Verifica o tratamento de OSError durante a exportacao."""

    caminho = tmp_path / "relatorio.csv"

    with patch.object(
        pd.DataFrame,
        "to_csv",
        side_effect=OSError("falha no disco"),
    ):
        with pytest.raises(OSError, match="Erro de sistema"):
            exportar_csv(dataframe_valido, str(caminho))


def test_exportar_csv_erro_inesperado(dataframe_valido, tmp_path):
    """Verifica o tratamento de erros inesperados."""

    caminho = tmp_path / "relatorio.csv"

    with patch.object(
        pd.DataFrame,
        "to_csv",
        side_effect=RuntimeError("falha inesperada"),
    ):
        with pytest.raises(RuntimeError, match="Erro inesperado"):
            exportar_csv(dataframe_valido, str(caminho))


# ----------------------------------------------------------------------------
# TESTES DA FUNCAO exportar_parquet()
# ----------------------------------------------------------------------------


def test_exportar_parquet_chama_to_parquet_corretamente(
    dataframe_valido,
    tmp_path,
):
    """Verifica os argumentos padrao utilizados para Parquet."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet") as mock_to_parquet:
        resultado = exportar_parquet(dataframe_valido, str(caminho))

    mock_to_parquet.assert_called_once_with(
        str(caminho.resolve()),
        index=False,
        engine="pyarrow",
        compression="snappy",
    )
    assert resultado is None


@pytest.mark.parametrize(
    "compressao",
    ["snappy", "gzip", "brotli", "lz4", "zstd", None],
)
def test_exportar_parquet_compressoes_validas(
    dataframe_valido,
    tmp_path,
    compressao,
):
    """Verifica todos os metodos de compressao permitidos."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet") as mock_to_parquet:
        exportar_parquet(
            dataframe_valido,
            str(caminho),
            compressao=compressao,
        )

    assert mock_to_parquet.call_args.kwargs["compression"] == compressao


def test_exportar_parquet_padroniza_compressao(
    dataframe_valido,
    tmp_path,
):
    """Verifica strip e lower no nome da compressao."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet") as mock_to_parquet:
        exportar_parquet(
            dataframe_valido,
            str(caminho),
            compressao="  GZIP  ",
        )

    assert mock_to_parquet.call_args.kwargs["compression"] == "gzip"


def test_exportar_parquet_cria_pasta_destino(dataframe_valido, tmp_path):
    """Verifica se a pasta de destino e criada automaticamente."""

    caminho = tmp_path / "nova_pasta" / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet"):
        exportar_parquet(dataframe_valido, str(caminho))

    assert caminho.parent.exists()
    assert caminho.parent.is_dir()


def test_exportar_parquet_exibe_mensagem_sucesso(
    dataframe_valido,
    tmp_path,
    capsys,
):
    """Verifica a mensagem exibida apos a exportacao."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet"):
        exportar_parquet(dataframe_valido, str(caminho))

    saida = capsys.readouterr().out
    assert "Arquivo Parquet exportado com sucesso" in saida
    assert str(caminho.resolve()) in saida


def test_exportar_parquet_nao_altera_dataframe_original(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a exportacao nao altera o DataFrame original."""

    original = dataframe_valido.copy(deep=True)
    caminho = tmp_path / "relatorio.parquet"

    with patch.object(pd.DataFrame, "to_parquet"):
        exportar_parquet(dataframe_valido, str(caminho))

    pd.testing.assert_frame_equal(dataframe_valido, original)


@pytest.mark.parametrize(
    "objeto_invalido",
    [None, [], {}, "texto", 10, 10.5],
)
def test_exportar_parquet_dataframe_invalido(objeto_invalido, tmp_path):
    """Verifica o erro quando o objeto não é um DataFrame."""

    caminho = tmp_path / "relatorio.parquet"

    with pytest.raises(TypeError, match="não é um DataFrame"):
        exportar_parquet(objeto_invalido, str(caminho))


def test_exportar_parquet_dataframe_vazio(dataframe_vazio, tmp_path):
    """Verifica o erro quando o DataFrame está vazio."""

    caminho = tmp_path / "relatorio.parquet"

    with pytest.raises(ValueError, match="DataFrame está vazio"):
        exportar_parquet(dataframe_vazio, str(caminho))


@pytest.mark.parametrize(
    "caminho_invalido",
    [None, 10, 10.5, [], {}, ("arquivo.parquet",)],
)
def test_exportar_parquet_caminho_nao_e_string(
    dataframe_valido,
    caminho_invalido,
):
    """Verifica o erro quando o caminho nao e uma string."""

    with pytest.raises(TypeError, match="'caminho'.*string"):
        exportar_parquet(dataframe_valido, caminho_invalido)


@pytest.mark.parametrize("caminho_vazio", ["", " ", "   ", "\t", "\n"])
def test_exportar_parquet_caminho_vazio(dataframe_valido, caminho_vazio):
    """Verifica o erro quando o caminho esta vazio."""

    with pytest.raises(ValueError, match="'caminho'.*vazio"):
        exportar_parquet(dataframe_valido, caminho_vazio)


@pytest.mark.parametrize(
    "extensao_invalida",
    ["relatorio.xlsx", "relatorio.csv", "relatorio.txt", "relatorio"],
)
def test_exportar_parquet_extensao_invalida(
    dataframe_valido,
    tmp_path,
    extensao_invalida,
):
    """Verifica o erro quando a extensao nao e Parquet."""

    caminho = tmp_path / extensao_invalida

    with pytest.raises(ValueError, match="extensão '.parquet'"):
        exportar_parquet(dataframe_valido, str(caminho))


def test_exportar_parquet_aceita_extensao_maiuscula(
    dataframe_valido,
    tmp_path,
):
    """Verifica se a extensao PARQUET em maiusculo e aceita."""

    caminho = tmp_path / "relatorio.PARQUET"

    with patch.object(pd.DataFrame, "to_parquet") as mock_to_parquet:
        exportar_parquet(dataframe_valido, str(caminho))

    assert mock_to_parquet.call_count == 1


@pytest.mark.parametrize("compressao_invalida", [10, 10.5, [], {}, ("gzip",)])
def test_exportar_parquet_compressao_tipo_invalido(
    dataframe_valido,
    tmp_path,
    compressao_invalida,
):
    """Verifica o erro quando compressao nao e string nem None."""

    caminho = tmp_path / "relatorio.parquet"

    with pytest.raises(TypeError, match="'compressao'.*string ou None"):
        exportar_parquet(
            dataframe_valido,
            str(caminho),
            compressao=compressao_invalida,
        )


@pytest.mark.parametrize(
    "compressao_invalida",
    ["", " ", "zip", "rar", "deflate", "invalida"],
)
def test_exportar_parquet_compressao_invalida(
    dataframe_valido,
    tmp_path,
    compressao_invalida,
):
    """Verifica o erro quando a compressao nao e permitida."""

    caminho = tmp_path / "relatorio.parquet"

    with pytest.raises(ValueError, match="compressão inválido"):
        exportar_parquet(
            dataframe_valido,
            str(caminho),
            compressao=compressao_invalida,
        )


def test_exportar_parquet_erro_ao_criar_pasta(dataframe_valido, tmp_path):
    """Verifica o tratamento de erro na criacao da pasta."""

    caminho = tmp_path / "pasta" / "relatorio.parquet"

    with patch("exportacao.os.makedirs", side_effect=OSError("falha")):
        with pytest.raises(OSError, match="criar a pasta de destino"):
            exportar_parquet(dataframe_valido, str(caminho))


def test_exportar_parquet_erro_de_permissao(dataframe_valido, tmp_path):
    """Verifica o tratamento de PermissionError."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(
        pd.DataFrame,
        "to_parquet",
        side_effect=PermissionError("arquivo bloqueado"),
    ):
        with pytest.raises(PermissionError, match="permissão de escrita"):
            exportar_parquet(dataframe_valido, str(caminho))


def test_exportar_parquet_pyarrow_ausente(dataframe_valido, tmp_path):
    """Verifica o tratamento da ausencia do pyarrow."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(
        pd.DataFrame,
        "to_parquet",
        side_effect=ImportError("pyarrow ausente"),
    ):
        with pytest.raises(ModuleNotFoundError, match="pyarrow"):
            exportar_parquet(dataframe_valido, str(caminho))


def test_exportar_parquet_erro_de_sistema(dataframe_valido, tmp_path):
    """Verifica o tratamento de OSError durante a exportacao."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(
        pd.DataFrame,
        "to_parquet",
        side_effect=OSError("falha no disco"),
    ):
        with pytest.raises(OSError, match="Erro de sistema"):
            exportar_parquet(dataframe_valido, str(caminho))


def test_exportar_parquet_erro_inesperado(dataframe_valido, tmp_path):
    """Verifica o tratamento de erros inesperados."""

    caminho = tmp_path / "relatorio.parquet"

    with patch.object(
        pd.DataFrame,
        "to_parquet",
        side_effect=RuntimeError("falha inesperada"),
    ):
        with pytest.raises(RuntimeError, match="Erro inesperado"):
            exportar_parquet(dataframe_valido, str(caminho))
