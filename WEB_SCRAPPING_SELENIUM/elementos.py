"""
============================================================
Módulo: SELENIUM / elementos.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para localização, consulta e
    validação de elementos em páginas web utilizando Selenium.
Funções disponíveis:
    - buscar_elemento()
    - buscar_elementos()
    - existe_elemento()
    - elemento_esta_visivel()
    - elemento_esta_habilitado()
    - elemento_esta_selecionado()
    - obter_texto()
    - obter_textos()
    - obter_atributo()
    - obter_propriedade()
    - obter_valor_campo()
    - obter_tag_elemento()
    - obter_html_elemento()
    - contar_elementos()
    - obter_elemento_pai()
    - buscar_elemento_filho()
    - buscar_elementos_filhos()
Dependências:
    - selenium
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Implementação da busca de elementos únicos e múltiplos.
        - Implementação de verificações de existência e estado.
        - Implementação da leitura de textos, valores e atributos.
        - Implementação da busca de elementos pais e filhos.
        - Inclusão de validações de driver, localizador e timeout.
        - Inclusão de exceções personalizadas do módulo.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa exceções utilizadas no tratamento das buscas
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchElementException,
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)

# Importa as condições esperadas utilizadas nas esperas explícitas
from selenium.webdriver.support import expected_conditions as EC

# Importa os tipos principais utilizados pelo módulo
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

# Importa o componente de espera explícita do Selenium
from selenium.webdriver.support.ui import WebDriverWait


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------

# Define o formato de um localizador Selenium, por exemplo (By.ID, "campo")
Localizador = tuple[str, str]


# ----------------------------------------------------------------------------
# EXCEÇÕES PERSONALIZADAS
# ----------------------------------------------------------------------------

class ElementosError(Exception):
    """Erro base para operações relacionadas a elementos Selenium."""


class DriverInativoError(ElementosError):
    """O WebDriver informado não possui uma sessão ativa."""


class LocalizadorInvalidoError(ElementosError):
    """O localizador informado não possui um formato válido."""


class ElementoNaoEncontradoError(ElementosError):
    """O elemento não foi encontrado dentro do tempo limite."""


class ElementoObsoletoError(ElementosError):
    """O elemento não está mais associado ao DOM atual."""


# ----------------------------------------------------------------------------
# FUNÇÕES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------

def _validar_driver(driver: WebDriver) -> None:
    """
    Valida se o objeto informado possui uma sessão WebDriver ativa.

    Args:
        driver (WebDriver):
            Instância do Selenium WebDriver.

    Raises:
        TypeError:
            Caso o objeto informado não seja um WebDriver.
        DriverInativoError:
            Caso a sessão do WebDriver esteja encerrada ou inválida.
    """
    # Verifica se o objeto informado é uma instância de WebDriver
    if not isinstance(driver, WebDriver):
        raise TypeError(
            "O parâmetro 'driver' deve ser uma instância de WebDriver."
        )

    # Tenta acessar a sessão atual para confirmar que o driver está ativo
    try:
        if not driver.session_id:
            raise DriverInativoError(
                "O WebDriver informado não possui uma sessão ativa."
            )

        driver.current_window_handle

    except DriverInativoError:
        raise

    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ) as error:
        raise DriverInativoError(
            "O WebDriver informado não possui uma sessão ativa."
        ) from error


def _validar_elemento(elemento: WebElement) -> None:
    """
    Valida se o objeto informado é um WebElement.

    Args:
        elemento (WebElement):
            Elemento Selenium que será validado.

    Raises:
        TypeError:
            Caso o objeto informado não seja um WebElement.
        ElementoObsoletoError:
            Caso o elemento não esteja mais associado ao DOM.
    """
    # Verifica se o objeto informado é um WebElement
    if not isinstance(elemento, WebElement):
        raise TypeError(
            "O parâmetro 'elemento' deve ser uma instância de WebElement."
        )

    # Acessa uma propriedade do elemento para detectar referência obsoleta
    try:
        elemento.tag_name
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento informado não está mais associado ao DOM atual."
        ) from error


def _validar_localizador(localizador: Localizador) -> None:
    """
    Valida o formato de um localizador Selenium.

    Args:
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.

    Raises:
        LocalizadorInvalidoError:
            Caso o localizador não seja uma tupla com duas strings válidas.
    """
    # Verifica se o localizador é uma tupla com exatamente dois itens
    if not isinstance(localizador, tuple) or len(localizador) != 2:
        raise LocalizadorInvalidoError(
            "O parâmetro 'localizador' deve ser uma tupla com dois itens, "
            "por exemplo: (By.ID, 'nome_do_elemento')."
        )

    estrategia, valor = localizador

    # Verifica se a estratégia e o valor são strings
    if not isinstance(estrategia, str) or not isinstance(valor, str):
        raise LocalizadorInvalidoError(
            "A estratégia e o valor do localizador devem ser strings."
        )

    # Verifica se a estratégia e o valor possuem conteúdo
    if not estrategia.strip() or not valor.strip():
        raise LocalizadorInvalidoError(
            "A estratégia e o valor do localizador não podem estar vazios."
        )


def _validar_timeout(timeout: int | float) -> None:
    """
    Valida o tempo máximo utilizado nas esperas explícitas.

    Args:
        timeout (int | float):
            Tempo máximo da espera, em segundos.

    Raises:
        TypeError:
            Caso o timeout não seja numérico.
        ValueError:
            Caso o timeout seja menor ou igual a zero.
    """
    # Rejeita booleanos e valores que não sejam números
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)):
        raise TypeError(
            "O parâmetro 'timeout' deve ser numérico."
        )

    # Verifica se o timeout é maior que zero
    if timeout <= 0:
        raise ValueError(
            "O parâmetro 'timeout' deve ser maior que zero."
        )


def _normalizar_texto(texto: str, remover_espacos: bool) -> str:
    """
    Normaliza o texto retornado por um elemento.

    Args:
        texto (str):
            Texto que será tratado.
        remover_espacos (bool):
            Define se espaços no início e no final serão removidos.

    Returns:
        str:
            Texto original ou texto sem espaços nas extremidades.
    """
    if remover_espacos:
        return texto.strip()

    return texto


# ----------------------------------------------------------------------------
# FUNÇÕES DE LOCALIZAÇÃO
# ----------------------------------------------------------------------------

def buscar_elemento(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20,
    visivel: bool = False
) -> WebElement:
    """
    Busca um elemento após aguardar sua presença ou visibilidade.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo de espera, em segundos.
        visivel (bool):
            Quando True, aguarda o elemento ficar visível.

    Returns:
        WebElement:
            Elemento localizado na página.

    Raises:
        TypeError:
            Caso driver, timeout ou visivel possuam tipos inválidos.
        LocalizadorInvalidoError:
            Caso o localizador seja inválido.
        DriverInativoError:
            Caso o WebDriver não esteja ativo.
        ElementoNaoEncontradoError:
            Caso o elemento não seja encontrado no tempo definido.
    """
    _validar_driver(driver)
    _validar_localizador(localizador)
    _validar_timeout(timeout)

    if not isinstance(visivel, bool):
        raise TypeError(
            "O parâmetro 'visivel' deve ser booleano."
        )

    try:
        espera = WebDriverWait(driver, timeout)

        if visivel:
            return espera.until(
                EC.visibility_of_element_located(localizador)
            )

        return espera.until(
            EC.presence_of_element_located(localizador)
        )

    except TimeoutException as error:
        estado = "visível" if visivel else "presente no DOM"
        raise ElementoNaoEncontradoError(
            f"O elemento {localizador} não ficou {estado} "
            f"dentro de {timeout} segundo(s)."
        ) from error


def buscar_elementos(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20,
    aguardar: bool = True
) -> list[WebElement]:
    """
    Busca todos os elementos correspondentes ao localizador.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo de espera, em segundos.
        aguardar (bool):
            Quando True, aguarda ao menos um elemento ficar presente.

    Returns:
        list[WebElement]:
            Lista de elementos encontrados.

    Raises:
        ElementoNaoEncontradoError:
            Caso aguardar seja True e nenhum elemento seja encontrado.
    """
    _validar_driver(driver)
    _validar_localizador(localizador)
    _validar_timeout(timeout)

    if not isinstance(aguardar, bool):
        raise TypeError(
            "O parâmetro 'aguardar' deve ser booleano."
        )

    if not aguardar:
        return driver.find_elements(*localizador)

    try:
        elementos = WebDriverWait(driver, timeout).until(
            EC.presence_of_all_elements_located(localizador)
        )
        return list(elementos)

    except TimeoutException as error:
        raise ElementoNaoEncontradoError(
            f"Nenhum elemento foi encontrado para {localizador} "
            f"dentro de {timeout} segundo(s)."
        ) from error


def existe_elemento(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 3,
    visivel: bool = False
) -> bool:
    """
    Verifica se um elemento existe no DOM ou está visível na página.

    A função utiliza buscar_elemento() e converte apenas a ausência do
    elemento em False. Erros de configuração, localizador ou driver
    continuam sendo propagados para não esconder falhas da automação.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo de espera, em segundos.
        visivel (bool):
            Quando True, exige que o elemento esteja visível. Quando False,
            verifica somente sua presença no DOM.

    Returns:
        bool:
            True quando o elemento satisfizer a condição e False quando
            não for encontrado dentro do tempo definido.

    Raises:
        TypeError:
            Caso driver, timeout ou visivel possuam tipos inválidos.
        LocalizadorInvalidoError:
            Caso o localizador informado seja inválido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
    """
    try:
        buscar_elemento(
            driver=driver,
            localizador=localizador,
            timeout=timeout,
            visivel=visivel,
        )
        return True
    except ElementoNaoEncontradoError:
        return False


def contar_elementos(
    driver: WebDriver,
    localizador: Localizador
) -> int:
    """
    Retorna a quantidade de elementos existentes no DOM.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.

    Returns:
        int:
            Quantidade de elementos encontrados.
    """
    _validar_driver(driver)
    _validar_localizador(localizador)

    return len(driver.find_elements(*localizador))


# ----------------------------------------------------------------------------
# FUNÇÕES DE ESTADO
# ----------------------------------------------------------------------------

def elemento_esta_visivel(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 3
) -> bool:
    """
    Verifica se um elemento está presente e visível na página.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para o elemento ficar visível.

    Returns:
        bool:
            True quando o elemento estiver visível e False caso contrário.

    Raises:
        TypeError:
            Caso driver ou timeout possuam tipos inválidos.
        LocalizadorInvalidoError:
            Caso o localizador informado seja inválido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
    """
    return existe_elemento(
        driver=driver,
        localizador=localizador,
        timeout=timeout,
        visivel=True,
    )


def elemento_esta_habilitado(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 3
) -> bool:
    """
    Verifica se um elemento está habilitado para interação.

    A função primeiro confirma a existência do elemento e, em seguida,
    consulta o estado retornado pelo método is_enabled() do Selenium.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar o elemento.

    Returns:
        bool:
            True quando o elemento estiver habilitado e False quando estiver
            desabilitado, ausente ou ficar obsoleto durante a consulta.

    Raises:
        TypeError:
            Caso driver ou timeout possuam tipos inválidos.
        LocalizadorInvalidoError:
            Caso o localizador informado seja inválido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
    """
    if not existe_elemento(driver, localizador, timeout):
        return False

    try:
        elemento = buscar_elemento(driver, localizador, timeout)
        return elemento.is_enabled()
    except StaleElementReferenceException:
        return False


def elemento_esta_selecionado(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 3
) -> bool:
    """
    Verifica se um checkbox, radio button ou option está selecionado.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar o elemento.

    Returns:
        bool:
            True quando o elemento estiver selecionado e False quando estiver
            desmarcado, ausente ou ficar obsoleto durante a consulta.

    Raises:
        TypeError:
            Caso driver ou timeout possuam tipos inválidos.
        LocalizadorInvalidoError:
            Caso o localizador informado seja inválido.
        DriverInativoError:
            Caso o WebDriver não possua uma sessão ativa.
    """
    if not existe_elemento(driver, localizador, timeout):
        return False

    try:
        elemento = buscar_elemento(driver, localizador, timeout)
        return elemento.is_selected()
    except StaleElementReferenceException:
        return False


# ----------------------------------------------------------------------------
# FUNÇÕES DE LEITURA
# ----------------------------------------------------------------------------

def obter_texto(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20,
    remover_espacos: bool = True
) -> str:
    """
    Retorna o texto visível de um elemento.

    A função aguarda o elemento ficar visível antes de acessar sua
    propriedade text.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para o elemento ficar visível.
        remover_espacos (bool):
            Define se espaços nas extremidades do texto serão removidos.

    Returns:
        str:
            Texto visível do elemento.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ElementoNaoEncontradoError:
            Caso o elemento não fique visível no tempo definido.
        ElementoObsoletoError:
            Caso o elemento fique obsoleto durante a leitura.
    """
    if not isinstance(remover_espacos, bool):
        raise TypeError(
            "O parâmetro 'remover_espacos' deve ser booleano."
        )

    elemento = buscar_elemento(
        driver=driver,
        localizador=localizador,
        timeout=timeout,
        visivel=True,
    )

    try:
        return _normalizar_texto(elemento.text, remover_espacos)
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento ficou obsoleto durante a leitura do texto."
        ) from error


def obter_textos(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20,
    remover_espacos: bool = True,
    ignorar_vazios: bool = False
) -> list[str]:
    """
    Retorna os textos de todos os elementos encontrados.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar os elementos.
        remover_espacos (bool):
            Define se espaços nas extremidades serão removidos.
        ignorar_vazios (bool):
            Define se textos vazios serão removidos do resultado.

    Returns:
        list[str]:
            Lista de textos na mesma ordem dos elementos encontrados.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ElementoNaoEncontradoError:
            Caso nenhum elemento seja encontrado no tempo definido.
        ElementoObsoletoError:
            Caso algum elemento fique obsoleto durante a leitura.
    """
    if not isinstance(remover_espacos, bool):
        raise TypeError(
            "O parâmetro 'remover_espacos' deve ser booleano."
        )

    if not isinstance(ignorar_vazios, bool):
        raise TypeError(
            "O parâmetro 'ignorar_vazios' deve ser booleano."
        )

    elementos = buscar_elementos(driver, localizador, timeout)
    textos = [
        _normalizar_texto(elemento.text, remover_espacos)
        for elemento in elementos
    ]

    if ignorar_vazios:
        textos = [texto for texto in textos if texto]

    return textos


def obter_atributo(
    driver: WebDriver,
    localizador: Localizador,
    atributo: str,
    timeout: int | float = 20
) -> str | None:
    """
    Retorna o valor de um atributo HTML do elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        atributo (str):
            Nome do atributo que será consultado.
        timeout (int | float):
            Tempo máximo para localizar o elemento.

    Returns:
        str | None:
            Valor do atributo ou None quando o atributo não existir.

    Raises:
        TypeError:
            Caso atributo ou demais parâmetros possuam tipos inválidos.
        ValueError:
            Caso o nome do atributo esteja vazio.
        ElementoNaoEncontradoError:
            Caso o elemento não seja encontrado.
        ElementoObsoletoError:
            Caso o elemento fique obsoleto durante a leitura.
    """
    if not isinstance(atributo, str):
        raise TypeError(
            "O parâmetro 'atributo' deve ser uma string."
        )

    if not atributo.strip():
        raise ValueError(
            "O parâmetro 'atributo' não pode estar vazio."
        )

    elemento = buscar_elemento(driver, localizador, timeout)

    try:
        return elemento.get_attribute(atributo.strip())
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento ficou obsoleto durante a leitura do atributo."
        ) from error


def obter_propriedade(
    driver: WebDriver,
    localizador: Localizador,
    propriedade: str,
    timeout: int | float = 20
) -> object:
    """
    Retorna o valor de uma propriedade JavaScript do elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        propriedade (str):
            Nome da propriedade JavaScript que será consultada.
        timeout (int | float):
            Tempo máximo para localizar o elemento.

    Returns:
        object:
            Valor da propriedade consultada. O tipo depende da propriedade.

    Raises:
        TypeError:
            Caso propriedade ou demais parâmetros possuam tipos inválidos.
        ValueError:
            Caso o nome da propriedade esteja vazio.
        ElementoNaoEncontradoError:
            Caso o elemento não seja encontrado.
        ElementoObsoletoError:
            Caso o elemento fique obsoleto durante a leitura.
    """
    if not isinstance(propriedade, str):
        raise TypeError(
            "O parâmetro 'propriedade' deve ser uma string."
        )

    if not propriedade.strip():
        raise ValueError(
            "O parâmetro 'propriedade' não pode estar vazio."
        )

    elemento = buscar_elemento(driver, localizador, timeout)

    try:
        return elemento.get_property(propriedade.strip())
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento ficou obsoleto durante a leitura da propriedade."
        ) from error


def obter_valor_campo(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20
) -> str:
    """
    Retorna o valor atual de um campo pelo atributo value.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar o campo.

    Returns:
        str:
            Valor atual do campo ou uma string vazia quando value for None.

    Raises:
        ElementoNaoEncontradoError:
            Caso o campo não seja encontrado.
        ElementoObsoletoError:
            Caso o campo fique obsoleto durante a leitura.
    """
    valor = obter_atributo(
        driver=driver,
        localizador=localizador,
        atributo="value",
        timeout=timeout,
    )

    return "" if valor is None else valor


def obter_tag_elemento(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20
) -> str:
    """
    Retorna o nome da tag HTML de um elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar o elemento.

    Returns:
        str:
            Nome da tag HTML, como input, button, div ou table.

    Raises:
        ElementoNaoEncontradoError:
            Caso o elemento não seja encontrado.
        ElementoObsoletoError:
            Caso o elemento fique obsoleto durante a leitura.
    """
    elemento = buscar_elemento(driver, localizador, timeout)

    try:
        return elemento.tag_name
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento ficou obsoleto durante a leitura da tag."
        ) from error


def obter_html_elemento(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20,
    html_externo: bool = True
) -> str:
    """
    Retorna o HTML interno ou externo de um elemento.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor da localização.
        timeout (int | float):
            Tempo máximo para localizar o elemento.
        html_externo (bool):
            Quando True, retorna outerHTML. Quando False, retorna innerHTML.

    Returns:
        str:
            Conteúdo HTML solicitado ou uma string vazia quando inexistente.

    Raises:
        TypeError:
            Caso html_externo ou demais parâmetros possuam tipos inválidos.
        ElementoNaoEncontradoError:
            Caso o elemento não seja encontrado.
        ElementoObsoletoError:
            Caso o elemento fique obsoleto durante a leitura.
    """
    if not isinstance(html_externo, bool):
        raise TypeError(
            "O parâmetro 'html_externo' deve ser booleano."
        )

    atributo = "outerHTML" if html_externo else "innerHTML"
    html = obter_atributo(driver, localizador, atributo, timeout)

    return "" if html is None else html


# ----------------------------------------------------------------------------
# FUNÇÕES DE RELACIONAMENTO ENTRE ELEMENTOS
# ----------------------------------------------------------------------------

def obter_elemento_pai(
    driver: WebDriver,
    localizador: Localizador,
    timeout: int | float = 20
) -> WebElement:
    """
    Localiza um elemento e retorna seu elemento pai imediato.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor do elemento filho.
        timeout (int | float):
            Tempo máximo para localizar o elemento filho.

    Returns:
        WebElement:
            Elemento pai imediato no DOM.

    Raises:
        ElementoNaoEncontradoError:
            Caso o elemento original ou seu pai não seja encontrado.
    """
    elemento = buscar_elemento(driver, localizador, timeout)

    try:
        return elemento.find_element("xpath", "..")
    except (NoSuchElementException, StaleElementReferenceException) as error:
        raise ElementoNaoEncontradoError(
            f"Não foi possível localizar o elemento pai de {localizador}."
        ) from error


def buscar_elemento_filho(
    elemento_pai: WebElement,
    localizador: Localizador
) -> WebElement:
    """
    Busca um elemento filho dentro de outro WebElement.

    Args:
        elemento_pai (WebElement):
            Elemento que delimita o contexto da busca.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor do elemento filho.

    Returns:
        WebElement:
            Primeiro elemento filho correspondente ao localizador.

    Raises:
        TypeError:
            Caso elemento_pai não seja um WebElement.
        LocalizadorInvalidoError:
            Caso o localizador seja inválido.
        ElementoNaoEncontradoError:
            Caso o elemento filho não seja encontrado.
        ElementoObsoletoError:
            Caso o elemento pai não esteja mais associado ao DOM.
    """
    _validar_elemento(elemento_pai)
    _validar_localizador(localizador)

    try:
        return elemento_pai.find_element(*localizador)
    except NoSuchElementException as error:
        raise ElementoNaoEncontradoError(
            f"O elemento filho {localizador} não foi encontrado."
        ) from error
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento pai ficou obsoleto durante a busca."
        ) from error


def buscar_elementos_filhos(
    elemento_pai: WebElement,
    localizador: Localizador
) -> list[WebElement]:
    """
    Busca todos os elementos filhos dentro de outro WebElement.

    Args:
        elemento_pai (WebElement):
            Elemento que delimita o contexto da busca.
        localizador (Localizador):
            Tupla contendo a estratégia e o valor dos elementos filhos.

    Returns:
        list[WebElement]:
            Lista de elementos filhos. A lista será vazia quando não houver
            correspondências.

    Raises:
        TypeError:
            Caso elemento_pai não seja um WebElement.
        LocalizadorInvalidoError:
            Caso o localizador seja inválido.
        ElementoObsoletoError:
            Caso o elemento pai não esteja mais associado ao DOM.
    """
    _validar_elemento(elemento_pai)
    _validar_localizador(localizador)

    try:
        return elemento_pai.find_elements(*localizador)
    except StaleElementReferenceException as error:
        raise ElementoObsoletoError(
            "O elemento pai ficou obsoleto durante a busca."
        ) from error
