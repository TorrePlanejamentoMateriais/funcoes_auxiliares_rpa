"""
============================================================
Modulo: TESTES / test_navegacao.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Testes unitarios do modulo navegacao.py utilizando pytest e
    unittest.mock, sem inicializar navegadores reais.
Dependencias:
    - pytest
    - selenium
    - excecoes
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial dos testes unitarios.
        - Cobertura das validacoes e manipulacoes de URL.
        - Cobertura da navegacao, consultas e esperas explicitas.
        - Cobertura do fluxo controlado de navegacao.
============================================================
"""

from unittest.mock import MagicMock, PropertyMock, patch

import pytest
from selenium.common.exceptions import (
    InvalidSessionIdException,
    TimeoutException,
    WebDriverException,
)

import navegacao
from excecoes import (
    CarregamentoPaginaError,
    DriverInativoError,
    NavegacaoError,
    ParametroInvalidoError,
    TempoEsperaExcedidoError,
    UrlInvalidaError,
)
from navegacao import (
    _comparar_texto,
    _executar_com_tratamento,
    _executar_espera_navegacao,
    _validar_booleano,
    _validar_driver,
    _validar_texto,
    _validar_timeout,
    adicionar_parametros_url,
    avancar_pagina,
    aguardar_carregamento_pagina,
    aguardar_titulo_conter,
    aguardar_titulo_igual,
    aguardar_url_conter,
    aguardar_url_corresponder,
    aguardar_url_igual,
    aguardar_url_mudar,
    atualizar_pagina,
    atualizar_pagina_sem_cache,
    executar_navegacao,
    montar_url,
    navegar_para,
    navegar_para_url_base,
    normalizar_url,
    obter_codigo_fonte,
    obter_dominio_atual,
    obter_dominio_url,
    obter_parametros_url,
    obter_titulo_pagina,
    obter_url_atual,
    remover_parametros_url,
    titulo_atual_contem,
    titulo_atual_igual,
    url_atual_contem,
    url_atual_igual,
    validar_url,
    voltar_pagina,
)


@pytest.fixture
def driver_mock():
    """Cria um mock reutilizavel de WebDriver."""
    driver = MagicMock()
    driver.session_id = "sessao-teste"
    driver.current_window_handle = "janela-principal"
    driver.current_url = "https://exemplo.com/inicio"
    driver.title = "Pagina inicial"
    driver.page_source = "<html></html>"
    return driver


# ----------------------------------------------------------------------------
# VALIDACOES INTERNAS
# ----------------------------------------------------------------------------
def test_validar_driver_ativo(monkeypatch, driver_mock):
    monkeypatch.setattr(navegacao, "WebDriver", MagicMock)
    assert _validar_driver(driver_mock) is None


def test_validar_driver_tipo_invalido():
    with pytest.raises(TypeError, match="WebDriver"):
        _validar_driver("invalido")


def test_validar_driver_sem_sessao(monkeypatch, driver_mock):
    monkeypatch.setattr(navegacao, "WebDriver", MagicMock)
    driver_mock.session_id = None
    with pytest.raises(DriverInativoError):
        _validar_driver(driver_mock)


def test_validar_driver_sessao_invalida(monkeypatch, driver_mock):
    monkeypatch.setattr(navegacao, "WebDriver", MagicMock)
    type(driver_mock).current_window_handle = PropertyMock(
        side_effect=InvalidSessionIdException()
    )
    with pytest.raises(DriverInativoError):
        _validar_driver(driver_mock)


@pytest.mark.parametrize("valor", [1, 0.1, 30, 60.5])
def test_validar_timeout_valido(valor):
    assert _validar_timeout(valor) is None


@pytest.mark.parametrize("valor", [True, "10", None, []])
def test_validar_timeout_tipo_invalido(valor):
    with pytest.raises(TypeError):
        _validar_timeout(valor)


@pytest.mark.parametrize("valor", [0, -1, -0.1])
def test_validar_timeout_valor_invalido(valor):
    with pytest.raises(ValueError):
        _validar_timeout(valor)


def test_validar_texto_normaliza():
    assert _validar_texto("  teste  ", "texto") == "teste"


@pytest.mark.parametrize("valor", [None, 1, [], {}])
def test_validar_texto_tipo_invalido(valor):
    with pytest.raises(TypeError):
        _validar_texto(valor, "texto")


@pytest.mark.parametrize("valor", ["", "   ", "\n"])
def test_validar_texto_vazio(valor):
    with pytest.raises(ValueError):
        _validar_texto(valor, "texto")


@pytest.mark.parametrize("valor", [True, False])
def test_validar_booleano_valido(valor):
    assert _validar_booleano(valor, "opcao") is None


@pytest.mark.parametrize("valor", [1, 0, "sim", None])
def test_validar_booleano_invalido(valor):
    with pytest.raises(TypeError):
        _validar_booleano(valor, "opcao")


@pytest.mark.parametrize(
    "atual,esperado,exata,maiusculas,resultado",
    [
        ("Relatorio", "relatorio", True, False, True),
        ("Relatorio mensal", "mensal", False, False, True),
        ("Relatorio", "relatorio", True, True, False),
        ("Relatorio", "Relatorio", True, True, True),
        ("Inicio", "fim", False, False, False),
    ],
)
def test_comparar_texto(atual, esperado, exata, maiusculas, resultado):
    assert _comparar_texto(atual, esperado, exata, maiusculas) is resultado


@patch("navegacao._validar_driver")
def test_executar_com_tratamento_retorna_resultado(mock_validar, driver_mock):
    assert _executar_com_tratamento(driver_mock, lambda: "ok", "erro") == "ok"


@patch("navegacao._validar_driver")
def test_executar_com_tratamento_converte_erro(mock_validar, driver_mock):
    operacao = MagicMock(side_effect=WebDriverException())
    with pytest.raises(NavegacaoError):
        _executar_com_tratamento(driver_mock, operacao, "Falha")


# ----------------------------------------------------------------------------
# URLS
# ----------------------------------------------------------------------------
@pytest.mark.parametrize(
    "url",
    ["https://exemplo.com", "http://exemplo.com/pagina", "https://localhost:8080"],
)
def test_validar_url_absoluta_valida(url):
    assert validar_url(url) == url


def test_validar_url_relativa_quando_permitida():
    assert validar_url("/pedidos/1", exigir_protocolo=False) == "/pedidos/1"


@pytest.mark.parametrize("url", ["exemplo.com", "/pagina"])
def test_validar_url_exige_protocolo(url):
    with pytest.raises(UrlInvalidaError):
        validar_url(url)


def test_validar_url_rejeita_protocolo():
    with pytest.raises(UrlInvalidaError):
        validar_url("ftp://exemplo.com")


def test_validar_url_rejeita_dominio_ausente():
    with pytest.raises(UrlInvalidaError):
        validar_url("https:///pagina")


def test_normalizar_url_adiciona_https():
    assert normalizar_url("exemplo.com/pagina") == "https://exemplo.com/pagina"


def test_normalizar_url_remove_barra_final():
    assert normalizar_url(
        "https://exemplo.com/pagina/", remover_barra_final=True
    ) == "https://exemplo.com/pagina"


def test_montar_url_combina_base_e_caminho():
    assert montar_url("https://exemplo.com", "pedidos/1") == (
        "https://exemplo.com/pedidos/1"
    )


def test_obter_dominio_url_com_e_sem_porta():
    url = "https://exemplo.com:8080/pagina"
    assert obter_dominio_url(url, incluir_porta=True) == "exemplo.com:8080"
    assert obter_dominio_url(url, incluir_porta=False) == "exemplo.com"


def test_obter_parametros_url():
    resultado = obter_parametros_url("/pedidos?status=aberto&centro=1&centro=2")
    assert resultado == {"status": ["aberto"], "centro": ["1", "2"]}


def test_adicionar_parametros_url_sobrescreve():
    resultado = adicionar_parametros_url(
        "https://exemplo.com?p=1", {"p": 2, "status": "aberto"}
    )
    assert obter_parametros_url(resultado) == {"p": ["2"], "status": ["aberto"]}


def test_adicionar_parametros_url_combina_valores():
    resultado = adicionar_parametros_url(
        "https://exemplo.com?centro=1", {"centro": [2, 3]}, sobrescrever=False
    )
    assert obter_parametros_url(resultado) == {"centro": ["1", "2", "3"]}


def test_adicionar_parametros_url_rejeita_parametros_invalidos():
    with pytest.raises(TypeError):
        adicionar_parametros_url("https://exemplo.com", ["a"])


def test_remover_parametros_url_especificos():
    resultado = remover_parametros_url(
        "https://exemplo.com?a=1&b=2", ["a"]
    )
    assert resultado == "https://exemplo.com?b=2"


def test_remover_todos_parametros_url():
    assert remover_parametros_url("https://exemplo.com?a=1") == "https://exemplo.com"


# ----------------------------------------------------------------------------
# NAVEGACAO
# ----------------------------------------------------------------------------
@patch("navegacao.obter_url_atual", return_value="https://exemplo.com/destino")
@patch("navegacao.aguardar_carregamento_pagina", return_value=True)
@patch("navegacao._executar_com_tratamento")
def test_navegar_para(mock_executar, mock_aguardar, mock_url, driver_mock):
    resultado = navegar_para(driver_mock, "https://exemplo.com/destino", timeout=10)
    assert resultado == "https://exemplo.com/destino"
    operacao = mock_executar.call_args.args[1]
    operacao()
    driver_mock.get.assert_called_once_with("https://exemplo.com/destino")
    mock_aguardar.assert_called_once_with(driver_mock, 10)


@patch("navegacao.navegar_para", return_value="https://exemplo.com/pedidos")
def test_navegar_para_url_base(mock_navegar, driver_mock):
    resultado = navegar_para_url_base(driver_mock, "https://exemplo.com", "pedidos")
    assert resultado == "https://exemplo.com/pedidos"
    mock_navegar.assert_called_once()


@pytest.mark.parametrize(
    "funcao,metodo",
    [(voltar_pagina, "back"), (avancar_pagina, "forward"), (atualizar_pagina, "refresh")],
)
@patch("navegacao.obter_url_atual", return_value="https://exemplo.com")
@patch("navegacao.aguardar_carregamento_pagina", return_value=True)
@patch("navegacao._executar_com_tratamento")
def test_comandos_historico(
    mock_executar, mock_aguardar, mock_url, funcao, metodo, driver_mock
):
    assert funcao(driver_mock, timeout=5) == "https://exemplo.com"
    operacao = mock_executar.call_args.args[1]
    operacao()
    getattr(driver_mock, metodo).assert_called_once_with()
    mock_aguardar.assert_called_once_with(driver_mock, 5)


@patch("navegacao.obter_url_atual", return_value="https://exemplo.com")
@patch("navegacao.aguardar_carregamento_pagina", return_value=True)
@patch("navegacao._validar_driver")
def test_atualizar_pagina_sem_cache(mock_validar, mock_aguardar, mock_url, driver_mock):
    assert atualizar_pagina_sem_cache(driver_mock, 10) == "https://exemplo.com"
    driver_mock.execute_cdp_cmd.assert_called_once_with(
        "Page.reload", {"ignoreCache": True}
    )


@patch("navegacao._validar_driver")
def test_atualizar_pagina_sem_cache_trata_erro(mock_validar, driver_mock):
    driver_mock.execute_cdp_cmd.side_effect = WebDriverException()
    with pytest.raises(NavegacaoError):
        atualizar_pagina_sem_cache(driver_mock)


# ----------------------------------------------------------------------------
# CONSULTAS E COMPARACOES
# ----------------------------------------------------------------------------
@patch("navegacao._validar_driver")
def test_obter_url_atual(mock_validar, driver_mock):
    assert obter_url_atual(driver_mock) == "https://exemplo.com/inicio"


@patch("navegacao._validar_driver")
def test_obter_titulo_pagina(mock_validar, driver_mock):
    assert obter_titulo_pagina(driver_mock) == "Pagina inicial"


@patch("navegacao._validar_driver")
def test_obter_codigo_fonte(mock_validar, driver_mock):
    assert obter_codigo_fonte(driver_mock) == "<html></html>"


@patch("navegacao.obter_url_atual", return_value="https://exemplo.com:8080/inicio")
def test_obter_dominio_atual(mock_url, driver_mock):
    assert obter_dominio_atual(driver_mock, incluir_porta=False) == "exemplo.com"


@pytest.mark.parametrize(
    "funcao,valor,esperado",
    [
        (url_atual_igual, "https://EXEMPLO.com/inicio", True),
        (url_atual_contem, "/inicio", True),
    ],
)
@patch("navegacao._validar_driver")
def test_comparacoes_url(mock_validar, funcao, valor, esperado, driver_mock):
    assert funcao(driver_mock, valor) is esperado


@pytest.mark.parametrize(
    "funcao,valor,esperado",
    [
        (titulo_atual_igual, "pagina inicial", True),
        (titulo_atual_contem, "inicial", True),
    ],
)
@patch("navegacao._validar_driver")
def test_comparacoes_titulo(mock_validar, funcao, valor, esperado, driver_mock):
    assert funcao(driver_mock, valor) is esperado


# ----------------------------------------------------------------------------
# ESPERAS
# ----------------------------------------------------------------------------
@patch("navegacao._validar_driver")
@patch("navegacao.WebDriverWait")
def test_executar_espera_navegacao_retorna_resultado(
    mock_wait, mock_validar, driver_mock
):
    condicao = MagicMock()
    mock_wait.return_value.until.return_value = "ok"
    assert _executar_espera_navegacao(driver_mock, condicao, 10, "erro") == "ok"
    mock_wait.assert_called_once_with(driver_mock, 10)


@patch("navegacao._validar_driver")
@patch("navegacao.WebDriverWait")
def test_executar_espera_navegacao_converte_timeout(
    mock_wait, mock_validar, driver_mock
):
    mock_wait.return_value.until.side_effect = TimeoutException()
    with pytest.raises(TempoEsperaExcedidoError):
        _executar_espera_navegacao(driver_mock, MagicMock(), 1, "timeout")


@pytest.mark.parametrize(
    "funcao,valor",
    [
        (aguardar_url_igual, "https://exemplo.com"),
        (aguardar_url_conter, "/inicio"),
        (aguardar_url_mudar, "https://exemplo.com/anterior"),
        (aguardar_url_corresponder, r"/pedido/\d+"),
        (aguardar_titulo_igual, "Pagina inicial"),
        (aguardar_titulo_conter, "inicial"),
    ],
)
@patch("navegacao._executar_espera_navegacao", return_value=True)
def test_wrappers_de_espera(mock_executar, funcao, valor, driver_mock):
    assert funcao(driver_mock, valor, timeout=10) is True
    mock_executar.assert_called_once()


@patch("navegacao._executar_espera_navegacao")
def test_aguardar_carregamento_pagina(mock_executar, driver_mock):
    mock_executar.side_effect = lambda driver, condicao, *args: condicao(driver)
    driver_mock.execute_script.return_value = "complete"
    assert aguardar_carregamento_pagina(driver_mock, 10) is True


@patch(
    "navegacao._executar_espera_navegacao",
    side_effect=TempoEsperaExcedidoError("timeout"),
)
def test_aguardar_carregamento_converte_erro(mock_executar, driver_mock):
    with pytest.raises(CarregamentoPaginaError):
        aguardar_carregamento_pagina(driver_mock)


# ----------------------------------------------------------------------------
# FLUXO CONTROLADO
# ----------------------------------------------------------------------------
@patch("navegacao.navegar_para")
@patch("navegacao.obter_url_atual", return_value="https://exemplo.com/origem")
def test_executar_navegacao_retorna_resultado_e_restaura(
    mock_url, mock_navegar, driver_mock
):
    acao = MagicMock(return_value={"status": "ok"})
    resultado = executar_navegacao(
        driver_mock,
        "https://exemplo.com/destino",
        acao,
        timeout=10,
        retornar_url_anterior=True,
    )
    assert resultado == {"status": "ok"}
    acao.assert_called_once_with(driver_mock)
    assert mock_navegar.call_count == 2
    assert mock_navegar.call_args_list[1].kwargs["url"] == "https://exemplo.com/origem"


def test_executar_navegacao_rejeita_acao_invalida(driver_mock):
    with pytest.raises(ParametroInvalidoError):
        executar_navegacao(driver_mock, "https://exemplo.com", "invalido")
