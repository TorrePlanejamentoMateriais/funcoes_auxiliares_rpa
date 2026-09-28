"""
============================================================
Módulo: SELENIUM / espera.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Funções auxiliares para esperas explícitas e sincronização
    de automações web utilizando Selenium.
Funções disponíveis:
    - criar_espera()
    - aguardar_condicao()
    - aguardar_elemento_presente()
    - aguardar_elementos_presentes()
    - aguardar_elemento_visivel()
    - aguardar_elementos_visiveis()
    - aguardar_elemento_clicavel()
    - aguardar_elemento_invisivel()
    - aguardar_elemento_desaparecer()
    - aguardar_elemento_habilitado()
    - aguardar_elemento_desabilitado()
    - aguardar_elemento_selecionado()
    - aguardar_elemento_nao_selecionado()
    - aguardar_elemento_obsoleto()
    - aguardar_texto_elemento()
    - aguardar_texto_valor_campo()
    - aguardar_texto_atributo()
    - aguardar_atributo_existir()
    - aguardar_atributo_igual()
    - aguardar_quantidade_elementos()
    - aguardar_quantidade_minima_elementos()
    - aguardar_quantidade_maxima_elementos()
    - aguardar_titulo_igual()
    - aguardar_titulo_conter()
    - aguardar_url_igual()
    - aguardar_url_conter()
    - aguardar_url_mudar()
    - aguardar_url_corresponder()
    - aguardar_alerta()
    - aguardar_frame_disponivel()
    - aguardar_nova_janela()
    - aguardar_quantidade_janelas()
    - aguardar_carregamento_pagina()
    - aguardar_javascript()
    - aguardar_spinner_desaparecer()
Dependências:
    - selenium
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão de esperas para elementos, textos e atributos.
        - Inclusão de esperas para URLs, janelas, frames e alertas.
        - Inclusão de validações e exceções personalizadas.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------
from collections.abc import Callable
from typing import Any, TypeVar

from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchElementException,
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------
Localizador = tuple[str, str]
T = TypeVar("T")


# ----------------------------------------------------------------------------
# EXCEÇÕES PERSONALIZADAS
# ----------------------------------------------------------------------------
class EsperaError(Exception):
    """Erro base para operações relacionadas às esperas Selenium."""


class DriverInativoError(EsperaError):
    """O WebDriver informado não possui uma sessão ativa."""


class LocalizadorInvalidoError(EsperaError):
    """O localizador informado não possui um formato válido."""


class CondicaoInvalidaError(EsperaError):
    """A condição personalizada informada não é válida."""


class TempoEsperaExcedidoError(EsperaError):
    """A condição não foi atendida dentro do tempo limite."""


# ----------------------------------------------------------------------------
# FUNÇÕES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------
def _validar_driver(driver: WebDriver) -> None:
    """
    Valida se o objeto informado possui uma sessão WebDriver ativa.

    Args:
        driver (WebDriver): Instância que será validada.

    Raises:
        TypeError: Caso o objeto não seja um WebDriver.
        DriverInativoError: Caso a sessão esteja inativa.
    """
    if not isinstance(driver, WebDriver):
        raise TypeError("O parâmetro 'driver' deve ser uma instância de WebDriver.")

    if not driver.session_id:
        raise DriverInativoError("O WebDriver não possui uma sessão ativa.")

    try:
        driver.current_window_handle
    except (InvalidSessionIdException, NoSuchWindowException, WebDriverException) as error:
        raise DriverInativoError("O WebDriver não possui uma sessão ativa.") from error


def _validar_elemento(elemento: WebElement) -> None:
    """
    Valida se o objeto informado é uma instância de WebElement.

    Args:
        elemento (WebElement):
            Elemento Selenium que será monitorado.

    Returns:
        None
                    A função apenas realiza a validação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if not isinstance(elemento, WebElement):
        raise TypeError("O parâmetro 'elemento' deve ser uma instância de WebElement.")


def _validar_localizador(localizador: Localizador) -> None:
    """
    Valida a estrutura e o conteúdo de um localizador Selenium.

    Args:
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.

    Returns:
        None
                    A função apenas realiza a validação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if not isinstance(localizador, tuple) or len(localizador) != 2:
        raise LocalizadorInvalidoError(
            "O localizador deve ser uma tupla como (By.ID, 'elemento')."
        )

    estrategia, valor = localizador
    if not all(isinstance(item, str) and item.strip() for item in (estrategia, valor)):
        raise LocalizadorInvalidoError(
            "A estratégia e o valor do localizador devem ser strings não vazias."
        )


def _validar_numero_positivo(valor: int | float, nome: str) -> None:
    """
    Valida um valor numérico que obrigatoriamente deve ser maior que zero.

    Args:
        valor (int | float):
            Valor numérico que será validado.
        nome (str):
            Nome do parâmetro utilizado nas mensagens de erro.

    Returns:
        None
                    A função apenas realiza a validação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise TypeError(f"O parâmetro '{nome}' deve ser numérico.")
    if valor <= 0:
        raise ValueError(f"O parâmetro '{nome}' deve ser maior que zero.")


def _validar_quantidade(quantidade: int) -> None:
    """
    Valida uma quantidade inteira maior ou igual a zero.

    Args:
        quantidade (int):
            Quantidade esperada de elementos ou janelas.

    Returns:
        None
                    A função apenas realiza a validação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if isinstance(quantidade, bool) or not isinstance(quantidade, int):
        raise TypeError("O parâmetro 'quantidade' deve ser inteiro.")
    if quantidade < 0:
        raise ValueError("O parâmetro 'quantidade' não pode ser negativo.")


def _validar_texto(texto: str, nome: str = "texto", vazio: bool = False) -> str:
    """
    Valida e normaliza um parâmetro textual.

    Args:
        texto (str):
            Texto que será aguardado ou validado.
        nome (str):
            Nome do parâmetro utilizado nas mensagens de erro.
        vazio (bool):
            Define se uma string vazia é permitida.

    Returns:
        str
                    Texto ou identificador retornado após a condição ser atendida.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if not isinstance(texto, str):
        raise TypeError(f"O parâmetro '{nome}' deve ser uma string.")
    resultado = texto.strip()
    if not vazio and not resultado:
        raise ValueError(f"O parâmetro '{nome}' não pode estar vazio.")
    return resultado


def _executar_espera(
    driver: WebDriver,
    condicao: Callable[[WebDriver], T],
    timeout: int | float,
    mensagem: str,
    frequencia: int | float = 0.5,
    excecoes_ignoradas: tuple[type[Exception], ...] | None = None,
) -> T:
    """
    Executa uma espera explícita e converte TimeoutException em erro próprio.

    Args:
        driver (WebDriver): Instância ativa do navegador.
        condicao (Callable): Condição verificada repetidamente.
        timeout (int | float): Tempo máximo da espera.
        mensagem (str): Mensagem utilizada em caso de timeout.
        frequencia (int | float): Intervalo entre verificações.
        excecoes_ignoradas (tuple | None): Exceções ignoradas nas tentativas.

    Returns:
        T: Resultado produzido pela condição.

    Raises:
        TempoEsperaExcedidoError: Caso a condição não seja atendida.
    """
    _validar_driver(driver)
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(frequencia, "frequencia")

    parametros = {"timeout": timeout, "poll_frequency": frequencia}
    if excecoes_ignoradas is not None:
        parametros["ignored_exceptions"] = excecoes_ignoradas

    try:
        return WebDriverWait(driver, **parametros).until(condicao)
    except TimeoutException as error:
        raise TempoEsperaExcedidoError(mensagem) from error


# ----------------------------------------------------------------------------
# FUNÇÕES GERAIS
# ----------------------------------------------------------------------------
def criar_espera(
    driver: WebDriver,
    timeout: int | float = 20,
    frequencia: int | float = 0.5,
    excecoes_ignoradas: tuple[type[Exception], ...] | None = None,
) -> WebDriverWait:
    """
    Cria uma instância configurada de WebDriverWait.

    Args:
        driver (WebDriver): Instância ativa do navegador.
        timeout (int | float): Tempo máximo da espera.
        frequencia (int | float): Intervalo entre verificações.
        excecoes_ignoradas (tuple | None): Classes de exceção ignoradas.

    Returns:
        WebDriverWait: Objeto de espera explícita configurado.
    """
    _validar_driver(driver)
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(frequencia, "frequencia")

    parametros = {"timeout": timeout, "poll_frequency": frequencia}
    if excecoes_ignoradas is not None:
        if not isinstance(excecoes_ignoradas, tuple) or not all(
            isinstance(item, type) and issubclass(item, Exception)
            for item in excecoes_ignoradas
        ):
            raise TypeError(
                "'excecoes_ignoradas' deve ser uma tupla de classes de exceção."
            )
        parametros["ignored_exceptions"] = excecoes_ignoradas

    return WebDriverWait(driver, **parametros)


def aguardar_condicao(
    driver: WebDriver,
    condicao: Callable[[WebDriver], T],
    timeout: int | float = 20,
    mensagem: str = "A condição esperada não foi atendida.",
    frequencia: int | float = 0.5,
) -> T:
    """
    Aguarda uma condição personalizada retornar um valor verdadeiro.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        condicao (Callable):
            Função ou objeto chamável avaliado pelo WebDriverWait.
        timeout (int | float):
            Tempo máximo da espera, em segundos.
        mensagem (str):
            Mensagem utilizada quando o tempo máximo for excedido.
        frequencia (int | float):
            Intervalo entre as verificações da condição, em segundos.

    Returns:
        T
                    Resultado produzido quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if not callable(condicao):
        raise CondicaoInvalidaError("O parâmetro 'condicao' deve ser chamável.")
    mensagem = _validar_texto(mensagem, "mensagem")
    return _executar_espera(driver, condicao, timeout, mensagem, frequencia)


# ----------------------------------------------------------------------------
# ESPERAS DE ELEMENTOS
# ----------------------------------------------------------------------------
def aguardar_elemento_presente(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um elemento existir no DOM, mesmo que ainda não esteja visível.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return _executar_espera(driver, EC.presence_of_element_located(localizador), timeout, f"O elemento {localizador} não ficou presente.")


def aguardar_elementos_presentes(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> list[WebElement]:
    """
    Aguarda pelo menos um elemento correspondente existir no DOM.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return list(_executar_espera(driver, EC.presence_of_all_elements_located(localizador), timeout, f"Nenhum elemento {localizador} ficou presente."))


def aguardar_elemento_visivel(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um elemento existir no DOM e ficar visível na página.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return _executar_espera(driver, EC.visibility_of_element_located(localizador), timeout, f"O elemento {localizador} não ficou visível.")


def aguardar_elementos_visiveis(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> list[WebElement]:
    """
    Aguarda todos os elementos localizados ficarem visíveis na página.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return list(_executar_espera(driver, EC.visibility_of_all_elements_located(localizador), timeout, f"Os elementos {localizador} não ficaram visíveis."))


def aguardar_elemento_clicavel(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um elemento ficar visível e habilitado para clique.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return _executar_espera(driver, EC.element_to_be_clickable(localizador), timeout, f"O elemento {localizador} não ficou clicável.")


def aguardar_elemento_invisivel(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> bool:
    """
    Aguarda um elemento ficar invisível ou ser removido do DOM.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return bool(_executar_espera(driver, EC.invisibility_of_element_located(localizador), timeout, f"O elemento {localizador} permaneceu visível."))


def aguardar_elemento_desaparecer(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> bool:
    """
    Aguarda um elemento desaparecer da tela ou ser removido do DOM.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return aguardar_elemento_invisivel(driver, localizador, timeout)


def aguardar_elemento_habilitado(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um elemento existir e ficar habilitado para interação.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    def condicao(navegador: WebDriver):
        elemento = navegador.find_element(*localizador)
        return elemento if elemento.is_enabled() else False
    return _executar_espera(driver, condicao, timeout, f"O elemento {localizador} não ficou habilitado.", excecoes_ignoradas=(NoSuchElementException, StaleElementReferenceException))


def aguardar_elemento_desabilitado(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um elemento existir e ficar desabilitado para interação.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    def condicao(navegador: WebDriver):
        elemento = navegador.find_element(*localizador)
        return elemento if not elemento.is_enabled() else False
    return _executar_espera(driver, condicao, timeout, f"O elemento {localizador} não ficou desabilitado.", excecoes_ignoradas=(NoSuchElementException, StaleElementReferenceException))


def aguardar_elemento_selecionado(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> bool:
    """
    Aguarda um checkbox, radio button ou option ficar selecionado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return bool(_executar_espera(driver, EC.element_located_to_be_selected(localizador), timeout, f"O elemento {localizador} não ficou selecionado."))


def aguardar_elemento_nao_selecionado(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> bool:
    """
    Aguarda um checkbox, radio button ou option ficar não selecionado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    condicao = EC.element_located_selection_state_to_be(localizador, False)
    return bool(_executar_espera(driver, condicao, timeout, f"O elemento {localizador} permaneceu selecionado."))


def aguardar_elemento_obsoleto(driver: WebDriver, elemento: WebElement, timeout: int | float = 20) -> bool:
    """
    Aguarda um WebElement deixar de estar associado ao DOM atual.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        elemento (WebElement):
            Elemento Selenium que será monitorado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_elemento(elemento)
    return bool(_executar_espera(driver, EC.staleness_of(elemento), timeout, "O elemento permaneceu associado ao DOM."))


# ----------------------------------------------------------------------------
# ESPERAS DE TEXTO E ATRIBUTOS
# ----------------------------------------------------------------------------
def aguardar_texto_elemento(driver: WebDriver, localizador: Localizador, texto: str, timeout: int | float = 20) -> bool:
    """
    Aguarda determinado texto aparecer no conteúdo visível de um elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        texto (str):
            Texto que será aguardado ou validado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    texto = _validar_texto(texto)
    return bool(_executar_espera(driver, EC.text_to_be_present_in_element(localizador, texto), timeout, f"O texto '{texto}' não apareceu em {localizador}."))


def aguardar_texto_valor_campo(driver: WebDriver, localizador: Localizador, texto: str, timeout: int | float = 20) -> bool:
    """
    Aguarda determinado texto aparecer no atributo value de um campo.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        texto (str):
            Texto que será aguardado ou validado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    texto = _validar_texto(texto)
    return bool(_executar_espera(driver, EC.text_to_be_present_in_element_value(localizador, texto), timeout, f"O valor '{texto}' não apareceu em {localizador}."))


def aguardar_texto_atributo(driver: WebDriver, localizador: Localizador, atributo: str, texto: str, timeout: int | float = 20) -> bool:
    """
    Aguarda determinado texto aparecer em um atributo do elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        atributo (str):
            Nome do atributo HTML que será consultado.
        texto (str):
            Texto que será aguardado ou validado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    atributo = _validar_texto(atributo, "atributo")
    texto = _validar_texto(texto)
    condicao = EC.text_to_be_present_in_element_attribute(localizador, atributo, texto)
    return bool(_executar_espera(driver, condicao, timeout, f"O texto '{texto}' não apareceu no atributo '{atributo}'."))


def aguardar_atributo_existir(driver: WebDriver, localizador: Localizador, atributo: str, timeout: int | float = 20) -> bool:
    """
    Aguarda um atributo ser incluído no elemento localizado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        atributo (str):
            Nome do atributo HTML que será consultado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    atributo = _validar_texto(atributo, "atributo")
    condicao = EC.element_attribute_to_include(localizador, atributo)
    return bool(_executar_espera(driver, condicao, timeout, f"O atributo '{atributo}' não apareceu em {localizador}."))


def aguardar_atributo_igual(driver: WebDriver, localizador: Localizador, atributo: str, valor_esperado: str, timeout: int | float = 20) -> WebElement:
    """
    Aguarda um atributo possuir exatamente o valor esperado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        atributo (str):
            Nome do atributo HTML que será consultado.
        valor_esperado (str):
            Valor exato esperado para o atributo.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        WebElement
                    Elemento que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    atributo = _validar_texto(atributo, "atributo")
    valor_esperado = _validar_texto(valor_esperado, "valor_esperado", vazio=True)
    def condicao(navegador: WebDriver):
        elemento = navegador.find_element(*localizador)
        return elemento if elemento.get_attribute(atributo) == valor_esperado else False
    return _executar_espera(driver, condicao, timeout, f"O atributo '{atributo}' não recebeu o valor esperado.", excecoes_ignoradas=(NoSuchElementException, StaleElementReferenceException))


# ----------------------------------------------------------------------------
# ESPERAS DE QUANTIDADE
# ----------------------------------------------------------------------------
def _aguardar_quantidade(driver: WebDriver, localizador: Localizador, quantidade: int, comparador: Callable[[int, int], bool], mensagem: str, timeout: int | float) -> list[WebElement]:
    """
    Executa uma espera interna baseada na quantidade de elementos localizados.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        quantidade (int):
            Quantidade esperada de elementos ou janelas.
        comparador (Callable[[int, int], bool]):
            Função usada para comparar as quantidades.
        mensagem (str):
            Mensagem utilizada quando o tempo máximo for excedido.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    _validar_quantidade(quantidade)
    def condicao(navegador: WebDriver):
        elementos = navegador.find_elements(*localizador)
        return elementos if comparador(len(elementos), quantidade) else False
    return list(_executar_espera(driver, condicao, timeout, mensagem))


def aguardar_quantidade_elementos(driver: WebDriver, localizador: Localizador, quantidade: int, timeout: int | float = 20) -> list[WebElement]:
    """
    Aguarda a quantidade de elementos ser exatamente igual à informada.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        quantidade (int):
            Quantidade esperada de elementos ou janelas.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return _aguardar_quantidade(driver, localizador, quantidade, lambda atual, esperado: atual == esperado, "A quantidade exata de elementos não foi atingida.", timeout)


def aguardar_quantidade_minima_elementos(driver: WebDriver, localizador: Localizador, quantidade: int, timeout: int | float = 20) -> list[WebElement]:
    """
    Aguarda existir pelo menos a quantidade informada de elementos.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        quantidade (int):
            Quantidade esperada de elementos ou janelas.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return _aguardar_quantidade(driver, localizador, quantidade, lambda atual, minimo: atual >= minimo, "A quantidade mínima de elementos não foi atingida.", timeout)


def aguardar_quantidade_maxima_elementos(driver: WebDriver, localizador: Localizador, quantidade: int, timeout: int | float = 20) -> list[WebElement]:
    """
    Aguarda existir no máximo a quantidade informada de elementos.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        quantidade (int):
            Quantidade esperada de elementos ou janelas.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[WebElement]
                    Lista de elementos que atendeu à condição esperada.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return _aguardar_quantidade(driver, localizador, quantidade, lambda atual, maximo: atual <= maximo, "A quantidade de elementos permaneceu acima do máximo.", timeout)


# ----------------------------------------------------------------------------
# ESPERAS DE TÍTULO E URL
# ----------------------------------------------------------------------------
def aguardar_titulo_igual(driver: WebDriver, titulo: str, timeout: int | float = 20) -> bool:
    """
    Aguarda o título da página ser exatamente igual ao valor informado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        titulo (str):
            Título exato esperado para a página.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    titulo = _validar_texto(titulo, "titulo")
    return bool(_executar_espera(driver, EC.title_is(titulo), timeout, f"O título não ficou igual a '{titulo}'."))


def aguardar_titulo_conter(driver: WebDriver, texto: str, timeout: int | float = 20) -> bool:
    """
    Aguarda o título da página conter determinado texto.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        texto (str):
            Texto que será aguardado ou validado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    texto = _validar_texto(texto)
    return bool(_executar_espera(driver, EC.title_contains(texto), timeout, f"O título não apresentou '{texto}'."))


def aguardar_url_igual(driver: WebDriver, url: str, timeout: int | float = 20) -> bool:
    """
    Aguarda a URL atual ser exatamente igual à URL informada.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        url (str):
            URL exata esperada.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    url = _validar_texto(url, "url")
    return bool(_executar_espera(driver, EC.url_to_be(url), timeout, f"A URL não ficou igual a '{url}'."))


def aguardar_url_conter(driver: WebDriver, texto: str, timeout: int | float = 20) -> bool:
    """
    Aguarda a URL atual conter determinado texto.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        texto (str):
            Texto que será aguardado ou validado.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    texto = _validar_texto(texto)
    return bool(_executar_espera(driver, EC.url_contains(texto), timeout, f"A URL não apresentou '{texto}'."))


def aguardar_url_mudar(driver: WebDriver, url_anterior: str, timeout: int | float = 20) -> bool:
    """
    Aguarda a URL atual ficar diferente da URL anterior.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        url_anterior (str):
            URL existente antes da ação que causará a navegação.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    url_anterior = _validar_texto(url_anterior, "url_anterior")
    return bool(_executar_espera(driver, EC.url_changes(url_anterior), timeout, "A URL não foi alterada."))


def aguardar_url_corresponder(driver: WebDriver, expressao_regular: str, timeout: int | float = 20) -> bool:
    """
    Aguarda a URL atual corresponder à expressão regular informada.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        expressao_regular (str):
            Expressão regular usada para validar a URL.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    expressao_regular = _validar_texto(expressao_regular, "expressao_regular")
    return bool(_executar_espera(driver, EC.url_matches(expressao_regular), timeout, "A URL não correspondeu à expressão regular."))


# ----------------------------------------------------------------------------
# ESPERAS DE ALERTA, FRAME E JANELAS
# ----------------------------------------------------------------------------
def aguardar_alerta(driver: WebDriver, timeout: int | float = 20) -> Any:
    """
    Aguarda um alerta nativo estar presente e retorna o objeto do alerta.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        Any
                    Resultado retornado pela condição, alerta ou JavaScript.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return _executar_espera(driver, EC.alert_is_present(), timeout, "Nenhum alerta foi apresentado.")


def aguardar_frame_disponivel(driver: WebDriver, localizador: Localizador, timeout: int | float = 20) -> bool:
    """
    Aguarda um frame ficar disponível e muda o contexto para ele.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_localizador(localizador)
    return bool(_executar_espera(driver, EC.frame_to_be_available_and_switch_to_it(localizador), timeout, f"O frame {localizador} não ficou disponível."))


def aguardar_nova_janela(driver: WebDriver, janelas_anteriores: list[str], timeout: int | float = 20) -> str:
    """
    Aguarda uma nova janela ou aba e retorna seu identificador.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        janelas_anteriores (list[str]):
            Identificadores das janelas abertas antes da ação.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        str
                    Texto ou identificador retornado após a condição ser atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    if not isinstance(janelas_anteriores, list) or not all(isinstance(item, str) for item in janelas_anteriores):
        raise TypeError("'janelas_anteriores' deve ser uma lista de strings.")
    _executar_espera(driver, EC.new_window_is_opened(janelas_anteriores), timeout, "Nenhuma nova janela foi aberta.")
    novas = [item for item in driver.window_handles if item not in janelas_anteriores]
    if not novas:
        raise TempoEsperaExcedidoError("O identificador da nova janela não foi localizado.")
    return novas[-1]


def aguardar_quantidade_janelas(driver: WebDriver, quantidade: int, timeout: int | float = 20) -> list[str]:
    """
    Aguarda a quantidade exata de janelas ou abas abertas.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        quantidade (int):
            Quantidade esperada de elementos ou janelas.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        list[str]
                    Lista com os identificadores das janelas abertas.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    _validar_quantidade(quantidade)
    _executar_espera(driver, EC.number_of_windows_to_be(quantidade), timeout, f"A quantidade de janelas não ficou igual a {quantidade}.")
    return list(driver.window_handles)


# ----------------------------------------------------------------------------
# ESPERAS DE CARREGAMENTO E JAVASCRIPT
# ----------------------------------------------------------------------------
def aguardar_carregamento_pagina(driver: WebDriver, timeout: int | float = 30) -> bool:
    """
    Aguarda o document.readyState da página ficar igual a complete.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    def condicao(navegador: WebDriver):
        return navegador.execute_script("return document.readyState") == "complete"
    return bool(_executar_espera(driver, condicao, timeout, "A página não concluiu o carregamento."))


def aguardar_javascript(driver: WebDriver, script: str, timeout: int | float = 20, *argumentos: Any) -> Any:
    """
    Aguarda um script JavaScript retornar um valor verdadeiro.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        script (str):
            Código JavaScript que será executado repetidamente.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        Any
                    Resultado retornado pela condição, alerta ou JavaScript.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    script = _validar_texto(script, "script")
    def condicao(navegador: WebDriver):
        return navegador.execute_script(script, *argumentos)
    return _executar_espera(driver, condicao, timeout, "O JavaScript não retornou um valor verdadeiro.")


def aguardar_spinner_desaparecer(driver: WebDriver, localizador: Localizador, timeout: int | float = 60) -> bool:
    """
    Aguarda um spinner ou indicador de carregamento desaparecer.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Returns:
        bool
                    True quando a condição for atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condição não seja atendida dentro do tempo definido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
    """
    return aguardar_elemento_invisivel(driver, localizador, timeout)
