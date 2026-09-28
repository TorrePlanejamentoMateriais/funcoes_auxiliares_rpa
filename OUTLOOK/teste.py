"""Testes unitarios do modulo utils.py."""

from datetime import date, datetime
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

import utils
from utils import (
    CARACTERES_INVALIDOS_ARQUIVO,
    FORMATO_DATA_OUTLOOK,
    CaminhoUtilError,
    ColecaoCOMError,
    PropriedadeCOMError,
    ValidacaoUtilError,
    caminho_unico,
    colecao_para_lista,
    comparar_textos,
    contar_colecao,
    definir_propriedade,
    escapar_filtro_outlook,
    formatar_data_outlook,
    garantir_diretorio,
    liberar_objeto_com,
    mascarar_texto,
    mesclar_dicionarios,
    normalizar_email,
    normalizar_emails,
    normalizar_lista_textos,
    obter_item_colecao,
    obter_propriedade,
    primeiro_ou_none,
    propriedade_existe,
    remover_acentos,
    sanitizar_nome_arquivo,
    validar_booleano,
    validar_caminho_arquivo,
    validar_data,
    validar_inteiro,
    validar_numero,
    validar_objeto,
    validar_texto,
)


@pytest.fixture
def objeto_mock():
    objeto = MagicMock()
    objeto.Subject = "Assunto de teste"
    return objeto


@pytest.fixture
def colecao_mock():
    colecao = MagicMock()
    colecao.Count = 3
    colecao.Item.side_effect = ["A", "B", "C"]
    return colecao


# Propriedades COM

def test_obter_propriedade_retorna_valor(objeto_mock):
    assert obter_propriedade(objeto_mock, "Subject") == "Assunto de teste"


def test_obter_propriedade_retorna_padrao_para_none():
    assert obter_propriedade(None, "Subject", padrao="") == ""


def test_obter_propriedade_retorna_padrao_em_erro():
    class Objeto:
        @property
        def Valor(self):
            raise RuntimeError("erro")

    assert obter_propriedade(Objeto(), "Valor", padrao="padrao") == "padrao"


def test_obter_propriedade_gera_erro_quando_solicitado():
    class Objeto:
        @property
        def Valor(self):
            raise RuntimeError("erro")

    with pytest.raises(PropriedadeCOMError):
        obter_propriedade(Objeto(), "Valor", gerar_erro=True)


def test_obter_propriedade_none_gera_erro_quando_solicitado():
    with pytest.raises(PropriedadeCOMError):
        obter_propriedade(None, "Valor", gerar_erro=True)


def test_definir_propriedade(objeto_mock):
    assert definir_propriedade(objeto_mock, "Subject", "Novo assunto") is objeto_mock
    assert objeto_mock.Subject == "Novo assunto"


def test_definir_propriedade_trata_erro():
    class Objeto:
        @property
        def Valor(self):
            return ""

        @Valor.setter
        def Valor(self, valor):
            raise RuntimeError("erro")

    with pytest.raises(PropriedadeCOMError):
        definir_propriedade(Objeto(), "Valor", "novo")


def test_propriedade_existe():
    class Objeto:
        Valor = 10

    assert propriedade_existe(Objeto(), "Valor") is True
    assert propriedade_existe(Objeto(), "Inexistente") is False
    assert propriedade_existe(None, "Valor") is False


# Validacoes

def test_validar_objeto_retorna_objeto():
    objeto = object()
    assert validar_objeto(objeto) is objeto


def test_validar_objeto_rejeita_none():
    with pytest.raises(ValidacaoUtilError):
        validar_objeto(None, "email")


def test_validar_texto_remove_espacos():
    assert validar_texto("  texto  ") == "texto"


def test_validar_texto_preserva_espacos_e_quebras():
    assert validar_texto(
        "\n\n",
        permitir_vazio=True,
        remover_espacos=False,
    ) == "\n\n"


def test_validar_texto_permite_vazio():
    assert validar_texto("   ", permitir_vazio=True) == ""


@pytest.mark.parametrize("valor", [None, 1, [], {}])
def test_validar_texto_rejeita_tipo(valor):
    with pytest.raises(TypeError):
        validar_texto(valor)


def test_validar_texto_rejeita_vazio():
    with pytest.raises(ValidacaoUtilError):
        validar_texto("   ")


@pytest.mark.parametrize("valor", [True, False])
def test_validar_booleano_aceita(valor):
    assert validar_booleano(valor) is valor


@pytest.mark.parametrize("valor", [1, 0, "sim", None])
def test_validar_booleano_rejeita(valor):
    with pytest.raises(TypeError):
        validar_booleano(valor)


@pytest.mark.parametrize("valor", [0, 1, 10])
def test_validar_inteiro_aceita(valor):
    assert validar_inteiro(valor, minimo=0, maximo=10) == valor


@pytest.mark.parametrize("valor", [True, 1.5, "1", None])
def test_validar_inteiro_rejeita_tipo(valor):
    with pytest.raises(TypeError):
        validar_inteiro(valor)


def test_validar_inteiro_rejeita_limites():
    with pytest.raises(ValidacaoUtilError):
        validar_inteiro(0, minimo=1)
    with pytest.raises(ValidacaoUtilError):
        validar_inteiro(11, maximo=10)


@pytest.mark.parametrize("valor", [0, 1, 1.5, -2.5])
def test_validar_numero_aceita(valor):
    assert validar_numero(valor) == valor


@pytest.mark.parametrize("valor", [True, "1", None])
def test_validar_numero_rejeita_tipo(valor):
    with pytest.raises(TypeError):
        validar_numero(valor)


def test_validar_numero_rejeita_limites():
    with pytest.raises(ValidacaoUtilError):
        validar_numero(-1, minimo=0)
    with pytest.raises(ValidacaoUtilError):
        validar_numero(11, maximo=10)


@pytest.mark.parametrize("valor", [date(2026, 9, 28), datetime(2026, 9, 28, 10)])
def test_validar_data_aceita(valor):
    assert validar_data(valor) is valor


def test_validar_data_permite_none():
    assert validar_data(None, permitir_none=True) is None


def test_validar_data_rejeita_tipo():
    with pytest.raises(TypeError):
        validar_data("28/09/2026")


# Emails e listas

def test_normalizar_email():
    assert normalizar_email(" USUARIO@EMPRESA.COM ") == "usuario@empresa.com"


def test_normalizar_email_permite_vazio():
    assert normalizar_email("   ", permitir_vazio=True) == ""


@pytest.mark.parametrize(
    "email",
    ["usuario", "usuario@", "@empresa.com", "a@@empresa.com", "a@empresa"],
)
def test_normalizar_email_rejeita_invalido(email):
    with pytest.raises(ValidacaoUtilError):
        normalizar_email(email)


def test_normalizar_lista_textos_de_string():
    assert normalizar_lista_textos("RPA; Critico, Fornecedor") == [
        "RPA",
        "Critico",
        "Fornecedor",
    ]


def test_normalizar_lista_textos_remove_duplicados_e_vazios():
    assert normalizar_lista_textos([" RPA ", "", "RPA", "Critico"]) == [
        "RPA",
        "Critico",
    ]


def test_normalizar_lista_textos_mantem_duplicados():
    assert normalizar_lista_textos(
        ["RPA", "RPA"],
        remover_duplicados=False,
    ) == ["RPA", "RPA"]


def test_normalizar_lista_textos_rejeita_sem_valores():
    with pytest.raises(ValidacaoUtilError):
        normalizar_lista_textos(["", "   "])


def test_normalizar_emails():
    assert normalizar_emails(
        "A@EMPRESA.COM; b@empresa.com; a@empresa.com"
    ) == ["a@empresa.com", "b@empresa.com"]


# Textos e filtros

@pytest.mark.parametrize(
    "atual,esperado,exata,maiusculas,resultado",
    [
        ("Retorno", "retorno", True, False, True),
        ("Retorno fornecedor", "fornecedor", False, False, True),
        ("Retorno", "retorno", True, True, False),
        ("Retorno", "Outro", False, False, False),
    ],
)
def test_comparar_textos(atual, esperado, exata, maiusculas, resultado):
    assert comparar_textos(atual, esperado, exata, maiusculas) is resultado


def test_escapar_filtro_outlook_preserva_espacos():
    assert escapar_filtro_outlook(" Fornecedor d'Avila ") == (
        " Fornecedor d''Avila "
    )


def test_formatar_data_outlook_datetime():
    assert formatar_data_outlook(datetime(2026, 9, 28, 10, 30)) == (
        "09/28/2026 10:30 AM"
    )


def test_formatar_data_outlook_date():
    assert formatar_data_outlook(date(2026, 9, 28)) == (
        "09/28/2026 12:00 AM"
    )


# Colecoes COM

def test_contar_colecao(colecao_mock):
    assert contar_colecao(colecao_mock) == 3


def test_contar_colecao_trata_erro(colecao_mock):
    colecao_mock.Count = "invalido"
    with pytest.raises(ColecaoCOMError):
        contar_colecao(colecao_mock)


def test_obter_item_colecao(colecao_mock):
    assert obter_item_colecao(colecao_mock, 1) == "A"
    colecao_mock.Item.assert_called_once_with(1)


def test_obter_item_colecao_rejeita_indice_fora(colecao_mock):
    with pytest.raises(ColecaoCOMError):
        obter_item_colecao(colecao_mock, 4)


def test_colecao_para_lista(colecao_mock):
    assert colecao_para_lista(colecao_mock) == ["A", "B", "C"]
    assert colecao_mock.Item.call_args_list == [call(1), call(2), call(3)]


def test_colecao_para_lista_aplica_limite():
    colecao = MagicMock()
    colecao.Count = 3
    colecao.Item.side_effect = ["A", "B"]
    assert colecao_para_lista(colecao, limite=2) == ["A", "B"]


def test_colecao_para_lista_aplica_transformacao():
    colecao = MagicMock()
    colecao.Count = 2
    colecao.Item.side_effect = [1, 2]
    assert colecao_para_lista(colecao, transformacao=lambda x: x * 10) == [10, 20]


def test_colecao_para_lista_ignora_erros():
    colecao = MagicMock()
    colecao.Count = 3
    colecao.Item.side_effect = ["A", RuntimeError("erro"), "C"]
    assert colecao_para_lista(colecao, ignorar_erros=True) == ["A", "C"]


def test_colecao_para_lista_converte_erro():
    colecao = MagicMock()
    colecao.Count = 1
    colecao.Item.side_effect = RuntimeError("erro")
    with pytest.raises(ColecaoCOMError):
        colecao_para_lista(colecao)


def test_primeiro_ou_none():
    assert primeiro_ou_none([1, 2]) == 1
    assert primeiro_ou_none([]) is None


def test_primeiro_ou_none_rejeita_nao_iteravel():
    with pytest.raises(TypeError):
        primeiro_ou_none(1)


# Texto e arquivos

def test_remover_acentos():
    assert remover_acentos("Ação, São Paulo, café") == "Acao, Sao Paulo, cafe"


def test_sanitizar_nome_arquivo():
    assert sanitizar_nome_arquivo(' Relatorio: Fornecedor/A?*.xlsx ') == (
        "Relatorio_ Fornecedor_A__.xlsx"
    )


def test_sanitizar_nome_arquivo_limita_tamanho():
    assert sanitizar_nome_arquivo("abcdefghij.txt", tamanho_maximo=5) == "abcde"


def test_sanitizar_nome_arquivo_rejeita_resultado_vazio():
    with pytest.raises(CaminhoUtilError):
        sanitizar_nome_arquivo("...")


def test_validar_caminho_arquivo(tmp_path):
    arquivo = tmp_path / "retorno.xlsx"
    arquivo.write_bytes(b"dados")
    assert validar_caminho_arquivo(
        arquivo,
        extensoes=["xlsx", ".xlsm"],
    ) == arquivo.resolve()


def test_validar_caminho_arquivo_rejeita_inexistente(tmp_path):
    with pytest.raises(FileNotFoundError):
        validar_caminho_arquivo(tmp_path / "inexistente.xlsx")


def test_validar_caminho_arquivo_rejeita_extensao(tmp_path):
    arquivo = tmp_path / "arquivo.txt"
    arquivo.write_text("dados", encoding="utf-8")
    with pytest.raises(CaminhoUtilError):
        validar_caminho_arquivo(arquivo, extensoes=["xlsx"])


def test_validar_caminho_arquivo_permite_inexistente(tmp_path):
    arquivo = tmp_path / "novo.xlsx"
    assert validar_caminho_arquivo(
        arquivo,
        deve_existir=False,
        extensoes=["xlsx"],
    ) == arquivo.resolve()


def test_garantir_diretorio(tmp_path):
    diretorio = tmp_path / "a" / "b" / "c"
    resultado = garantir_diretorio(diretorio)
    assert resultado == diretorio.resolve()
    assert resultado.is_dir()


def test_caminho_unico_quando_nao_existe(tmp_path):
    caminho = tmp_path / "arquivo.xlsx"
    assert caminho_unico(caminho) == caminho.resolve()


def test_caminho_unico_adiciona_sufixo(tmp_path):
    original = tmp_path / "arquivo.xlsx"
    original.write_bytes(b"dados")
    (tmp_path / "arquivo_1.xlsx").write_bytes(b"dados")
    assert caminho_unico(original) == (tmp_path / "arquivo_2.xlsx").resolve()


# Dicionarios e mascaramento

def test_mesclar_dicionarios():
    assert mesclar_dicionarios(
        {"status": "PENDENTE", "tentativa": 1},
        {"status": "ENVIADO", "data": "2026-09-28"},
    ) == {
        "status": "ENVIADO",
        "tentativa": 1,
        "data": "2026-09-28",
    }


def test_mesclar_dicionarios_ignora_none():
    assert mesclar_dicionarios(
        {"status": "ENVIADO"},
        {"status": None, "tentativa": 2},
        ignorar_none=True,
    ) == {"status": "ENVIADO", "tentativa": 2}


def test_mesclar_dicionarios_rejeita_nao_mapeamento():
    with pytest.raises(TypeError):
        mesclar_dicionarios({"a": 1}, [1, 2])


@pytest.mark.parametrize(
    "texto,inicio,fim,esperado",
    [
        ("usuario@empresa.com", 2, 2, "us***************om"),
        ("abcd", 2, 2, "****"),
        ("abcdef", 0, 2, "****ef"),
        ("abcdef", 2, 0, "ab****"),
        ("", 2, 2, ""),
    ],
)
def test_mascarar_texto(texto, inicio, fim, esperado):
    assert mascarar_texto(texto, inicio, fim) == esperado


def test_mascarar_texto_caractere_personalizado():
    assert mascarar_texto("abcdef", 1, 1, "#") == "a####f"


def test_mascarar_texto_rejeita_caractere_invalido():
    with pytest.raises(TypeError):
        mascarar_texto("abcdef", caractere="**")


def test_liberar_objeto_com():
    assert liberar_objeto_com(MagicMock()) is None
    assert liberar_objeto_com(None) is None


# Estrutura publica

def test_constantes():
    assert CARACTERES_INVALIDOS_ARQUIVO == '<>:"/\\|?*'
    assert FORMATO_DATA_OUTLOOK == "%m/%d/%Y %I:%M %p"


def test_all_contem_nomes_existentes():
    assert utils.__all__
    for nome in utils.__all__:
        assert hasattr(utils, nome)


def test_all_sem_duplicidades():
    assert len(utils.__all__) == len(set(utils.__all__))


def test_excecoes_possuem_docstring():
    classes = [
        utils.OutlookUtilsError,
        utils.ValidacaoUtilError,
        utils.PropriedadeCOMError,
        utils.ColecaoCOMError,
        utils.CaminhoUtilError,
    ]
    for classe in classes:
        assert classe.__doc__
