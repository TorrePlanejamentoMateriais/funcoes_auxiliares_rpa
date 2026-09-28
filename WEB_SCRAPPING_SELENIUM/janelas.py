"""
============================================================
Modulo: SELENIUM / janelas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para controle de janelas e
    abas em automacoes web utilizando Selenium.
Funcoes disponiveis:
    - obter_janela_atual()
    - listar_janelas()
    - contar_janelas()
    - janela_existe()
    - obter_indice_janela_atual()
    - obter_janela_por_indice()
    - trocar_para_janela()
    - trocar_para_janela_por_indice()
    - trocar_para_primeira_janela()
    - trocar_para_ultima_janela()
    - trocar_para_proxima_janela()
    - trocar_para_janela_anterior()
    - trocar_para_janela_por_titulo()
    - trocar_para_janela_por_url()
    - abrir_nova_aba()
    - abrir_nova_janela()
    - aguardar_nova_janela()
    - aguardar_quantidade_janelas()
    - executar_em_nova_janela()
    - fechar_janela_atual()
    - fechar_janela()
    - fechar_outras_janelas()
    - fechar_todas_exceto()
    - obter_informacoes_janelas()
Dependencias:
    - selenium
    - excecoes
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de navegacao por identificador e indice.
        - Inclusao de pesquisa por titulo e URL.
        - Inclusao de criacao e espera de novas janelas.
        - Inclusao de fechamento seguro de abas e janelas.
        - Inclusao de coleta de informacoes das janelas abertas.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa Callable para tipar funcoes executadas em uma nova janela
from collections.abc import Callable

# Importa Any e Literal para tipagem dos retornos e tipos de janela
from typing import Any, Literal, TypeVar

# Importa as excecoes nativas tratadas pelo modulo
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    TimeoutException,
    WebDriverException,
)

# Importa o tipo principal do Selenium WebDriver
from selenium.webdriver.remote.webdriver import WebDriver

# Importa as condicoes e a espera explicita utilizadas pelo modulo
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Importa as excecoes centralizadas da biblioteca
from excecoes import (
    DriverInativoError,
    JanelaError,
    JanelaNaoEncontradaError,
    ParametroInvalidoError,
    TempoEsperaExcedidoError,
)


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------

# Define os tipos aceitos ao criar uma nova janela pelo Selenium
TipoJanela = Literal["tab", "window"]

# Define um tipo generico para o retorno de funcoes personalizadas
T = TypeVar("T")


# ----------------------------------------------------------------------------
# FUNCOES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------

def _validar_driver(driver: WebDriver) -> None:
    """
    Valida se o objeto informado possui uma sessao WebDriver ativa.

    Args:
        driver (WebDriver):
            Instancia do Selenium WebDriver que sera validada.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o objeto informado nao seja um WebDriver.
        DriverInativoError:
            Caso a sessao esteja encerrada ou invalida.
    """
    # Verifica se o objeto informado e uma instancia de WebDriver
    if not isinstance(driver, WebDriver):
        raise TypeError(
            "O parametro 'driver' deve ser uma instancia de WebDriver."
        )

    # Interrompe a execucao quando o driver nao possui identificador de sessao
    if not driver.session_id:
        raise DriverInativoError(
            "O WebDriver informado nao possui uma sessao ativa."
        )

    # Tenta acessar os identificadores para confirmar que a sessao responde
    try:
        driver.window_handles
    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ) as error:
        raise DriverInativoError(
            "O WebDriver informado nao possui uma sessao ativa."
        ) from error


def _validar_identificador(identificador: str) -> str:
    """
    Valida e normaliza o identificador de uma janela.

    Args:
        identificador (str):
            Identificador unico da janela ou aba.

    Returns:
        str:
            Identificador sem espacos nas extremidades.

    Raises:
        TypeError:
            Caso o identificador nao seja uma string.
        ValueError:
            Caso o identificador esteja vazio.
    """
    # Verifica se o identificador e uma string
    if not isinstance(identificador, str):
        raise TypeError(
            "O parametro 'identificador' deve ser uma string."
        )

    # Remove espacos desnecessarios das extremidades
    identificador_normalizado = identificador.strip()

    # Interrompe a execucao quando o identificador esta vazio
    if not identificador_normalizado:
        raise ValueError(
            "O parametro 'identificador' nao pode estar vazio."
        )

    # Retorna o identificador normalizado
    return identificador_normalizado


def _validar_indice(indice: int) -> None:
    """
    Valida o indice utilizado para selecionar uma janela.

    Args:
        indice (int):
            Posicao da janela na lista de identificadores.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o indice nao seja um numero inteiro.
    """
    # Rejeita booleanos porque bool tambem e uma subclasse de int
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError(
            "O parametro 'indice' deve ser um numero inteiro."
        )


def _validar_timeout(timeout: int | float) -> None:
    """
    Valida o tempo maximo utilizado nas esperas de janela.

    Args:
        timeout (int | float):
            Tempo maximo da espera, em segundos.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o timeout nao seja numerico.
        ValueError:
            Caso o timeout seja menor ou igual a zero.
    """
    # Rejeita booleanos e objetos que nao sejam numeros
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)):
        raise TypeError(
            "O parametro 'timeout' deve ser numerico."
        )

    # Verifica se o timeout e maior que zero
    if timeout <= 0:
        raise ValueError(
            "O parametro 'timeout' deve ser maior que zero."
        )


def _validar_texto(texto: str, nome_parametro: str) -> str:
    """
    Valida e normaliza um parametro textual obrigatorio.

    Args:
        texto (str):
            Texto que sera validado.
        nome_parametro (str):
            Nome utilizado na mensagem de erro.

    Returns:
        str:
            Texto sem espacos nas extremidades.

    Raises:
        TypeError:
            Caso o texto nao seja uma string.
        ValueError:
            Caso o texto esteja vazio.
    """
    # Verifica se o valor recebido e uma string
    if not isinstance(texto, str):
        raise TypeError(
            f"O parametro '{nome_parametro}' deve ser uma string."
        )

    # Remove espacos das extremidades
    texto_normalizado = texto.strip()

    # Interrompe a execucao quando o texto esta vazio
    if not texto_normalizado:
        raise ValueError(
            f"O parametro '{nome_parametro}' nao pode estar vazio."
        )

    # Retorna o texto tratado
    return texto_normalizado


def _comparar_texto(
    valor_atual: str,
    valor_procurado: str,
    correspondencia_exata: bool,
    considerar_maiusculas: bool,
) -> bool:
    """
    Compara titulo ou URL conforme as opcoes informadas.

    Args:
        valor_atual (str):
            Texto obtido na janela atual.
        valor_procurado (str):
            Texto utilizado como criterio da pesquisa.
        correspondencia_exata (bool):
            Define se os valores devem ser exatamente iguais.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia letras maiusculas.

    Returns:
        bool:
            True quando os valores correspondem e False caso contrario.

    Raises:
        TypeError:
            Caso as opcoes de comparacao nao sejam booleanas.
    """
    # Valida as opcoes booleanas utilizadas na comparacao
    if not isinstance(correspondencia_exata, bool):
        raise TypeError(
            "O parametro 'correspondencia_exata' deve ser booleano."
        )

    if not isinstance(considerar_maiusculas, bool):
        raise TypeError(
            "O parametro 'considerar_maiusculas' deve ser booleano."
        )

    # Mantem ou normaliza as letras conforme a opcao informada
    atual = valor_atual if considerar_maiusculas else valor_atual.lower()
    procurado = (
        valor_procurado
        if considerar_maiusculas
        else valor_procurado.lower()
    )

    # Realiza comparacao exata ou parcial
    if correspondencia_exata:
        return atual == procurado

    return procurado in atual


# ----------------------------------------------------------------------------
# FUNCOES DE CONSULTA
# ----------------------------------------------------------------------------

def obter_janela_atual(driver: WebDriver) -> str:
    """
    Retorna o identificador da janela ou aba atualmente selecionada.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Identificador unico da janela atual.

    Raises:
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        JanelaNaoEncontradaError:
            Caso a janela atual nao esteja mais disponivel.
    """
    # Valida a sessao antes de consultar a janela atual
    _validar_driver(driver)

    try:
        # Retorna o identificador da janela que possui o foco do WebDriver
        return driver.current_window_handle
    except (NoSuchWindowException, WebDriverException) as error:
        raise JanelaNaoEncontradaError(
            "Nao foi possivel obter a janela atualmente selecionada."
        ) from error


def listar_janelas(driver: WebDriver) -> list[str]:
    """
    Retorna os identificadores de todas as janelas e abas abertas.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        list[str]:
            Copia da lista de identificadores disponiveis.

    Raises:
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
    """
    # Valida a sessao antes de consultar os identificadores
    _validar_driver(driver)

    # Retorna uma copia para evitar alteracoes na estrutura do Selenium
    return list(driver.window_handles)


def contar_janelas(driver: WebDriver) -> int:
    """
    Retorna a quantidade de janelas e abas abertas.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        int:
            Quantidade atual de identificadores de janela.
    """
    # Reutiliza a listagem validada e retorna seu tamanho
    return len(listar_janelas(driver))


def janela_existe(
    driver: WebDriver,
    identificador: str,
) -> bool:
    """
    Verifica se um identificador de janela ainda esta disponivel.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        identificador (str):
            Identificador que sera procurado.

    Returns:
        bool:
            True quando a janela existir e False caso contrario.
    """
    # Valida e normaliza o identificador procurado
    identificador_validado = _validar_identificador(identificador)

    # Verifica a presenca do identificador na lista atual
    return identificador_validado in listar_janelas(driver)


def obter_indice_janela_atual(driver: WebDriver) -> int:
    """
    Retorna o indice da janela atualmente selecionada.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        int:
            Posicao da janela atual na lista de identificadores.

    Raises:
        JanelaNaoEncontradaError:
            Caso o identificador atual nao esteja na lista de janelas.
    """
    # Obtem o identificador e a lista de janelas atuais
    atual = obter_janela_atual(driver)
    janelas = listar_janelas(driver)

    try:
        # Retorna a posicao da janela atual
        return janelas.index(atual)
    except ValueError as error:
        raise JanelaNaoEncontradaError(
            "A janela atual nao foi encontrada na lista de janelas abertas."
        ) from error


def obter_janela_por_indice(
    driver: WebDriver,
    indice: int,
) -> str:
    """
    Retorna o identificador localizado em determinado indice.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        indice (int):
            Posicao da janela na lista de identificadores.

    Returns:
        str:
            Identificador da janela encontrada.

    Raises:
        JanelaNaoEncontradaError:
            Caso o indice esteja fora do intervalo disponivel.
    """
    # Valida o indice antes de consultar a lista
    _validar_indice(indice)
    janelas = listar_janelas(driver)

    try:
        # Permite indices positivos e negativos conforme o padrao do Python
        return janelas[indice]
    except IndexError as error:
        raise JanelaNaoEncontradaError(
            f"Nao existe janela no indice {indice}. "
            f"Quantidade disponivel: {len(janelas)}."
        ) from error


# ----------------------------------------------------------------------------
# FUNCOES DE TROCA DE CONTEXTO
# ----------------------------------------------------------------------------

def trocar_para_janela(
    driver: WebDriver,
    identificador: str,
) -> str:
    """
    Transfere o contexto do WebDriver para uma janela especifica.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        identificador (str):
            Identificador da janela que recebera o foco.

    Returns:
        str:
            Identificador da janela selecionada.

    Raises:
        JanelaNaoEncontradaError:
            Caso o identificador nao esteja disponivel ou a troca falhe.
    """
    # Valida a sessao e o identificador solicitado
    _validar_driver(driver)
    identificador_validado = _validar_identificador(identificador)

    # Evita solicitar a troca para uma janela inexistente
    if identificador_validado not in listar_janelas(driver):
        raise JanelaNaoEncontradaError(
            f"A janela '{identificador_validado}' nao foi encontrada."
        )

    try:
        # Transfere o contexto do Selenium para a janela escolhida
        driver.switch_to.window(identificador_validado)
    except (NoSuchWindowException, WebDriverException) as error:
        raise JanelaNaoEncontradaError(
            f"Nao foi possivel trocar para a janela '{identificador_validado}'."
        ) from error

    # Retorna o identificador selecionado
    return identificador_validado


def trocar_para_janela_por_indice(
    driver: WebDriver,
    indice: int,
) -> str:
    """
    Transfere o contexto para a janela localizada em determinado indice.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        indice (int):
            Posicao da janela na lista de identificadores.

    Returns:
        str:
            Identificador da janela selecionada.

    Raises:
        JanelaNaoEncontradaError:
            Caso nao exista janela no indice informado.
    """
    # Localiza o identificador correspondente ao indice
    identificador = obter_janela_por_indice(driver, indice)

    # Reutiliza a troca segura por identificador
    return trocar_para_janela(driver, identificador)


def trocar_para_primeira_janela(driver: WebDriver) -> str:
    """
    Transfere o contexto para a primeira janela aberta.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Identificador da primeira janela.
    """
    # A primeira janela ocupa o indice zero
    return trocar_para_janela_por_indice(driver, 0)


def trocar_para_ultima_janela(driver: WebDriver) -> str:
    """
    Transfere o contexto para a ultima janela aberta.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Identificador da ultima janela.
    """
    # O indice -1 representa o ultimo item da lista
    return trocar_para_janela_por_indice(driver, -1)


def trocar_para_proxima_janela(
    driver: WebDriver,
    circular: bool = True,
) -> str:
    """
    Transfere o contexto para a janela seguinte.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        circular (bool):
            Retorna a primeira janela quando a atual for a ultima.

    Returns:
        str:
            Identificador da janela selecionada.

    Raises:
        ParametroInvalidoError:
            Caso circular seja False e a janela atual seja a ultima.
    """
    # Valida a opcao de navegacao circular
    if not isinstance(circular, bool):
        raise TypeError(
            "O parametro 'circular' deve ser booleano."
        )

    # Obtem a lista e o indice atual
    janelas = listar_janelas(driver)
    indice_atual = obter_indice_janela_atual(driver)
    proximo_indice = indice_atual + 1

    # Trata o fim da lista conforme a opcao circular
    if proximo_indice >= len(janelas):
        if not circular:
            raise ParametroInvalidoError(
                "A janela atual ja e a ultima janela aberta."
            )
        proximo_indice = 0

    # Troca para a janela calculada
    return trocar_para_janela(driver, janelas[proximo_indice])


def trocar_para_janela_anterior(
    driver: WebDriver,
    circular: bool = True,
) -> str:
    """
    Transfere o contexto para a janela anterior.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        circular (bool):
            Seleciona a ultima janela quando a atual for a primeira.

    Returns:
        str:
            Identificador da janela selecionada.

    Raises:
        ParametroInvalidoError:
            Caso circular seja False e a janela atual seja a primeira.
    """
    # Valida a opcao de navegacao circular
    if not isinstance(circular, bool):
        raise TypeError(
            "O parametro 'circular' deve ser booleano."
        )

    # Obtem a lista e calcula o indice anterior
    janelas = listar_janelas(driver)
    indice_atual = obter_indice_janela_atual(driver)
    indice_anterior = indice_atual - 1

    # Trata o inicio da lista conforme a opcao circular
    if indice_anterior < 0:
        if not circular:
            raise ParametroInvalidoError(
                "A janela atual ja e a primeira janela aberta."
            )
        indice_anterior = len(janelas) - 1

    # Troca para a janela calculada
    return trocar_para_janela(driver, janelas[indice_anterior])


def trocar_para_janela_por_titulo(
    driver: WebDriver,
    titulo: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    restaurar_janela_original: bool = True,
) -> str:
    """
    Procura uma janela pelo titulo e transfere o contexto para ela.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        titulo (str):
            Titulo completo ou parcial procurado.
        correspondencia_exata (bool):
            Define se o titulo deve corresponder exatamente.
        considerar_maiusculas (bool):
            Define se a pesquisa diferencia maiusculas de minusculas.
        restaurar_janela_original (bool):
            Restaura a janela inicial quando nenhuma correspondencia existe.

    Returns:
        str:
            Identificador da janela correspondente.

    Raises:
        JanelaNaoEncontradaError:
            Caso nenhuma janela possua o titulo procurado.
    """
    # Valida o titulo e a opcao de restauracao
    titulo_validado = _validar_texto(titulo, "titulo")

    if not isinstance(restaurar_janela_original, bool):
        raise TypeError(
            "O parametro 'restaurar_janela_original' deve ser booleano."
        )

    # Armazena a janela inicial antes de percorrer as demais
    janela_original = obter_janela_atual(driver)

    # Percorre todas as janelas disponiveis
    for identificador in listar_janelas(driver):
        try:
            trocar_para_janela(driver, identificador)

            if _comparar_texto(
                driver.title,
                titulo_validado,
                correspondencia_exata,
                considerar_maiusculas,
            ):
                return identificador
        except JanelaNaoEncontradaError:
            continue

    # Restaura a janela inicial quando solicitado e ainda disponivel
    if restaurar_janela_original and janela_existe(driver, janela_original):
        trocar_para_janela(driver, janela_original)

    # Informa que nenhuma janela corresponde ao titulo
    raise JanelaNaoEncontradaError(
        f"Nenhuma janela com o titulo '{titulo_validado}' foi encontrada."
    )


def trocar_para_janela_por_url(
    driver: WebDriver,
    url: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    restaurar_janela_original: bool = True,
) -> str:
    """
    Procura uma janela pela URL e transfere o contexto para ela.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str):
            URL completa ou parcial procurada.
        correspondencia_exata (bool):
            Define se a URL deve corresponder exatamente.
        considerar_maiusculas (bool):
            Define se a pesquisa diferencia maiusculas de minusculas.
        restaurar_janela_original (bool):
            Restaura a janela inicial quando nenhuma correspondencia existe.

    Returns:
        str:
            Identificador da janela correspondente.

    Raises:
        JanelaNaoEncontradaError:
            Caso nenhuma janela possua a URL procurada.
    """
    # Valida a URL e a opcao de restauracao
    url_validada = _validar_texto(url, "url")

    if not isinstance(restaurar_janela_original, bool):
        raise TypeError(
            "O parametro 'restaurar_janela_original' deve ser booleano."
        )

    # Armazena a janela inicial antes da pesquisa
    janela_original = obter_janela_atual(driver)

    # Percorre as janelas e compara suas URLs
    for identificador in listar_janelas(driver):
        try:
            trocar_para_janela(driver, identificador)

            if _comparar_texto(
                driver.current_url,
                url_validada,
                correspondencia_exata,
                considerar_maiusculas,
            ):
                return identificador
        except JanelaNaoEncontradaError:
            continue

    # Restaura a janela inicial quando solicitado
    if restaurar_janela_original and janela_existe(driver, janela_original):
        trocar_para_janela(driver, janela_original)

    # Informa que nenhuma URL correspondente foi encontrada
    raise JanelaNaoEncontradaError(
        f"Nenhuma janela com a URL '{url_validada}' foi encontrada."
    )


# ----------------------------------------------------------------------------
# FUNCOES DE CRIACAO E ESPERA
# ----------------------------------------------------------------------------

def abrir_nova_aba(
    driver: WebDriver,
    url: str | None = None,
) -> str:
    """
    Cria uma nova aba, transfere o contexto e opcionalmente acessa uma URL.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str | None):
            URL opcional que sera aberta na nova aba.

    Returns:
        str:
            Identificador da nova aba.

    Raises:
        JanelaError:
            Caso a nova aba nao possa ser criada.
    """
    # Valida a sessao antes de criar a aba
    _validar_driver(driver)

    # Valida a URL somente quando ela for informada
    url_validada = None

    if url is not None:
        url_validada = _validar_texto(url, "url")

    try:
        # O Selenium 4 cria a aba e transfere automaticamente o contexto
        driver.switch_to.new_window("tab")
        identificador = driver.current_window_handle

        # Abre a URL solicitada na nova aba
        if url_validada is not None:
            driver.get(url_validada)

        return identificador
    except WebDriverException as error:
        raise JanelaError(
            "Nao foi possivel criar uma nova aba."
        ) from error


def abrir_nova_janela(
    driver: WebDriver,
    url: str | None = None,
) -> str:
    """
    Cria uma nova janela, transfere o contexto e opcionalmente acessa uma URL.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str | None):
            URL opcional que sera aberta na nova janela.

    Returns:
        str:
            Identificador da nova janela.

    Raises:
        JanelaError:
            Caso a nova janela nao possa ser criada.
    """
    # Valida a sessao antes de criar a janela
    _validar_driver(driver)

    # Valida a URL somente quando ela for informada
    url_validada = None

    if url is not None:
        url_validada = _validar_texto(url, "url")

    try:
        # O Selenium 4 cria a janela e transfere automaticamente o contexto
        driver.switch_to.new_window("window")
        identificador = driver.current_window_handle

        # Abre a URL solicitada na nova janela
        if url_validada is not None:
            driver.get(url_validada)

        return identificador
    except WebDriverException as error:
        raise JanelaError(
            "Nao foi possivel criar uma nova janela."
        ) from error


def aguardar_nova_janela(
    driver: WebDriver,
    janelas_anteriores: list[str] | tuple[str, ...],
    timeout: int | float = 20,
    trocar_contexto: bool = True,
) -> str:
    """
    Aguarda uma nova janela ou aba e retorna seu identificador.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        janelas_anteriores (list[str] | tuple[str, ...]):
            Identificadores existentes antes da acao que abre a janela.
        timeout (int | float):
            Tempo maximo para a nova janela aparecer.
        trocar_contexto (bool):
            Define se o WebDriver sera transferido para a nova janela.

    Returns:
        str:
            Identificador da nova janela ou aba.

    Raises:
        TempoEsperaExcedidoError:
            Caso nenhuma nova janela seja aberta no tempo definido.
        TypeError:
            Caso a colecao de identificadores seja invalida.
    """
    # Valida a sessao e os parametros da espera
    _validar_driver(driver)
    _validar_timeout(timeout)

    if not isinstance(janelas_anteriores, (list, tuple)):
        raise TypeError(
            "O parametro 'janelas_anteriores' deve ser uma lista ou tupla."
        )

    if not all(isinstance(item, str) and item.strip() for item in janelas_anteriores):
        raise TypeError(
            "Todos os itens de 'janelas_anteriores' devem ser strings nao vazias."
        )

    if not isinstance(trocar_contexto, bool):
        raise TypeError(
            "O parametro 'trocar_contexto' deve ser booleano."
        )

    # Cria uma copia imutavel dos identificadores anteriores
    anteriores = list(janelas_anteriores)

    try:
        # Aguarda o Selenium detectar o aumento da quantidade de janelas
        WebDriverWait(driver, timeout).until(
            EC.new_window_is_opened(anteriores)
        )
    except TimeoutException as error:
        raise TempoEsperaExcedidoError(
            "Nenhuma nova janela ou aba foi aberta dentro do tempo limite."
        ) from error

    # Identifica os handles que nao existiam anteriormente
    novas = [
        identificador
        for identificador in listar_janelas(driver)
        if identificador not in anteriores
    ]

    if not novas:
        raise JanelaNaoEncontradaError(
            "A condicao de nova janela foi atendida, mas nenhum novo "
            "identificador foi localizado."
        )

    # Seleciona a janela mais recentemente adicionada
    nova_janela = novas[-1]

    # Transfere o contexto quando solicitado
    if trocar_contexto:
        trocar_para_janela(driver, nova_janela)

    # Retorna o novo identificador
    return nova_janela


def aguardar_quantidade_janelas(
    driver: WebDriver,
    quantidade: int,
    timeout: int | float = 20,
) -> list[str]:
    """
    Aguarda a quantidade exata de janelas ou abas abertas.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        quantidade (int):
            Quantidade exata de janelas esperada.
        timeout (int | float):
            Tempo maximo para a condicao ser atendida.

    Returns:
        list[str]:
            Identificadores disponiveis apos a espera.

    Raises:
        TempoEsperaExcedidoError:
            Caso a quantidade nao seja atingida no tempo definido.
    """
    # Valida a sessao, a quantidade e o timeout
    _validar_driver(driver)
    _validar_timeout(timeout)

    if isinstance(quantidade, bool) or not isinstance(quantidade, int):
        raise TypeError(
            "O parametro 'quantidade' deve ser um numero inteiro."
        )

    if quantidade <= 0:
        raise ValueError(
            "O parametro 'quantidade' deve ser maior que zero."
        )

    try:
        # Aguarda a quantidade exata por meio de Expected Conditions
        WebDriverWait(driver, timeout).until(
            EC.number_of_windows_to_be(quantidade)
        )
    except TimeoutException as error:
        raise TempoEsperaExcedidoError(
            f"A quantidade de janelas nao ficou igual a {quantidade}."
        ) from error

    # Retorna a lista atualizada dos identificadores
    return listar_janelas(driver)


def executar_em_nova_janela(
    driver: WebDriver,
    acao_abrir: Callable[[], Any],
    acao_executar: Callable[[WebDriver], T],
    timeout: int | float = 20,
    fechar_apos_execucao: bool = True,
    restaurar_janela_original: bool = True,
) -> T:
    """
    Abre uma nova janela, executa uma funcao nela e restaura o contexto.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        acao_abrir (Callable[[], Any]):
            Funcao que realiza a acao responsavel por abrir a nova janela.
        acao_executar (Callable[[WebDriver], T]):
            Funcao executada apos a troca para a nova janela.
        timeout (int | float):
            Tempo maximo para a nova janela aparecer.
        fechar_apos_execucao (bool):
            Define se a nova janela sera fechada ao final.
        restaurar_janela_original (bool):
            Define se o contexto inicial sera restaurado.

    Returns:
        T:
            Resultado retornado por acao_executar.

    Raises:
        ParametroInvalidoError:
            Caso uma das acoes nao seja chamavel.
        TempoEsperaExcedidoError:
            Caso a nova janela nao seja aberta no tempo definido.
    """
    # Valida as funcoes recebidas
    if not callable(acao_abrir):
        raise ParametroInvalidoError(
            "O parametro 'acao_abrir' deve ser uma funcao ou objeto chamavel."
        )

    if not callable(acao_executar):
        raise ParametroInvalidoError(
            "O parametro 'acao_executar' deve ser uma funcao ou objeto chamavel."
        )

    # Valida as opcoes booleanas
    if not isinstance(fechar_apos_execucao, bool):
        raise TypeError(
            "O parametro 'fechar_apos_execucao' deve ser booleano."
        )

    if not isinstance(restaurar_janela_original, bool):
        raise TypeError(
            "O parametro 'restaurar_janela_original' deve ser booleano."
        )

    # Registra o contexto antes de executar a acao de abertura
    janela_original = obter_janela_atual(driver)
    janelas_anteriores = listar_janelas(driver)

    # Executa a acao responsavel por abrir a janela
    acao_abrir()

    # Aguarda a nova janela e transfere o contexto para ela
    nova_janela = aguardar_nova_janela(
        driver=driver,
        janelas_anteriores=janelas_anteriores,
        timeout=timeout,
        trocar_contexto=True,
    )

    try:
        # Executa a funcao principal no contexto da nova janela
        return acao_executar(driver)
    finally:
        # Fecha a nova janela quando solicitado e quando ela ainda existe
        if fechar_apos_execucao and janela_existe(driver, nova_janela):
            fechar_janela(
                driver=driver,
                identificador=nova_janela,
                janela_destino=janela_original,
            )
        # Restaura a janela original quando solicitado e disponivel
        elif (
            restaurar_janela_original
            and janela_existe(driver, janela_original)
        ):
            trocar_para_janela(driver, janela_original)


# ----------------------------------------------------------------------------
# FUNCOES DE FECHAMENTO
# ----------------------------------------------------------------------------

def fechar_janela_atual(
    driver: WebDriver,
    janela_destino: str | None = None,
) -> str | None:
    """
    Fecha a janela atual e transfere o contexto para outra janela aberta.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        janela_destino (str | None):
            Janela que recebera o foco apos o fechamento. Quando None,
            utiliza a ultima janela restante.

    Returns:
        str | None:
            Identificador selecionado ou None quando nao restam janelas.

    Raises:
        JanelaNaoEncontradaError:
            Caso a janela de destino nao esteja disponivel.
        JanelaError:
            Caso a janela atual nao possa ser fechada.
    """
    # Valida a sessao e registra a janela atual
    _validar_driver(driver)
    janela_atual = obter_janela_atual(driver)

    # Valida o destino antes do fechamento quando ele foi informado
    destino_validado = None

    if janela_destino is not None:
        destino_validado = _validar_identificador(janela_destino)

        if destino_validado == janela_atual:
            raise ParametroInvalidoError(
                "A janela de destino nao pode ser a janela que sera fechada."
            )

        if not janela_existe(driver, destino_validado):
            raise JanelaNaoEncontradaError(
                f"A janela de destino '{destino_validado}' nao foi encontrada."
            )

    try:
        # Fecha somente a janela que possui o foco atual
        driver.close()
    except (NoSuchWindowException, WebDriverException) as error:
        raise JanelaError(
            f"Nao foi possivel fechar a janela '{janela_atual}'."
        ) from error

    # Consulta os identificadores restantes diretamente
    try:
        janelas_restantes = list(driver.window_handles)
    except (InvalidSessionIdException, WebDriverException):
        return None

    # Retorna None quando a ultima janela foi fechada
    if not janelas_restantes:
        return None

    # Seleciona o destino informado ou a ultima janela restante
    destino = destino_validado or janelas_restantes[-1]

    # Garante que o destino ainda esta disponivel
    if destino not in janelas_restantes:
        destino = janelas_restantes[-1]

    # Transfere o contexto e retorna o identificador selecionado
    driver.switch_to.window(destino)
    return destino


def fechar_janela(
    driver: WebDriver,
    identificador: str,
    janela_destino: str | None = None,
) -> str | None:
    """
    Fecha uma janela especifica e seleciona outra janela disponivel.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        identificador (str):
            Identificador da janela que sera fechada.
        janela_destino (str | None):
            Janela selecionada apos o fechamento.

    Returns:
        str | None:
            Identificador selecionado ou None quando nao restam janelas.

    Raises:
        JanelaNaoEncontradaError:
            Caso a janela que sera fechada nao exista.
    """
    # Valida e seleciona a janela que sera fechada
    identificador_validado = _validar_identificador(identificador)
    trocar_para_janela(driver, identificador_validado)

    # Fecha a janela agora selecionada
    return fechar_janela_atual(
        driver=driver,
        janela_destino=janela_destino,
    )


def fechar_outras_janelas(
    driver: WebDriver,
    manter_janela: str | None = None,
) -> str:
    """
    Fecha todas as janelas, exceto uma janela principal.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        manter_janela (str | None):
            Janela preservada. Quando None, utiliza a janela atual.

    Returns:
        str:
            Identificador da unica janela preservada.

    Raises:
        JanelaNaoEncontradaError:
            Caso a janela que deve ser mantida nao exista.
    """
    # Define a janela preservada conforme o argumento recebido
    janela_principal = (
        obter_janela_atual(driver)
        if manter_janela is None
        else _validar_identificador(manter_janela)
    )

    # Verifica se a janela principal ainda existe
    if not janela_existe(driver, janela_principal):
        raise JanelaNaoEncontradaError(
            f"A janela que deveria ser mantida nao existe: '{janela_principal}'."
        )

    # Percorre uma copia da lista para permitir o fechamento seguro
    for identificador in listar_janelas(driver):
        if identificador == janela_principal:
            continue

        trocar_para_janela(driver, identificador)
        driver.close()

    # Restaura o contexto para a janela preservada
    trocar_para_janela(driver, janela_principal)
    return janela_principal


def fechar_todas_exceto(
    driver: WebDriver,
    identificadores_manter: list[str] | tuple[str, ...] | set[str],
    janela_destino: str | None = None,
) -> list[str]:
    """
    Fecha todas as janelas que nao estiverem na colecao de preservacao.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        identificadores_manter (list | tuple | set):
            Identificadores que nao serao fechados.
        janela_destino (str | None):
            Janela que recebera o foco ao final.

    Returns:
        list[str]:
            Identificadores restantes apos os fechamentos.

    Raises:
        TypeError:
            Caso a colecao de identificadores possua formato invalido.
        JanelaNaoEncontradaError:
            Caso nenhum identificador preservado exista.
    """
    # Valida a colecao de janelas que serao preservadas
    if not isinstance(identificadores_manter, (list, tuple, set)):
        raise TypeError(
            "O parametro 'identificadores_manter' deve ser lista, tupla ou set."
        )

    if not identificadores_manter:
        raise ValueError(
            "O parametro 'identificadores_manter' nao pode estar vazio."
        )

    if not all(
        isinstance(item, str) and item.strip()
        for item in identificadores_manter
    ):
        raise TypeError(
            "Todos os identificadores mantidos devem ser strings nao vazias."
        )

    # Normaliza os identificadores e remove duplicidades
    manter = {
        item.strip()
        for item in identificadores_manter
    }
    janelas_atuais = listar_janelas(driver)
    existentes = manter.intersection(janelas_atuais)

    # Exige ao menos uma janela valida para preservar a sessao
    if not existentes:
        raise JanelaNaoEncontradaError(
            "Nenhuma das janelas informadas para preservacao esta aberta."
        )

    # Fecha as janelas que nao pertencem ao conjunto de preservacao
    for identificador in list(janelas_atuais):
        if identificador in existentes:
            continue

        trocar_para_janela(driver, identificador)
        driver.close()

    # Obtem as janelas restantes diretamente
    restantes = list(driver.window_handles)

    # Define a janela que recebera o foco final
    if janela_destino is not None:
        destino = _validar_identificador(janela_destino)

        if destino not in restantes:
            raise JanelaNaoEncontradaError(
                f"A janela de destino '{destino}' nao foi preservada."
            )
    else:
        # Preserva a ordem fornecida pelo Selenium
        destino = restantes[0]

    # Transfere o contexto para a janela final
    driver.switch_to.window(destino)
    return restantes


# ----------------------------------------------------------------------------
# FUNCOES DE INFORMACOES
# ----------------------------------------------------------------------------

def obter_informacoes_janelas(
    driver: WebDriver,
    restaurar_janela_original: bool = True,
    ignorar_janelas_indisponiveis: bool = True,
) -> list[dict[str, Any]]:
    """
    Retorna identificador, indice, titulo, URL e estado de cada janela.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        restaurar_janela_original (bool):
            Define se a janela inicialmente selecionada sera restaurada.
        ignorar_janelas_indisponiveis (bool):
            Ignora janelas fechadas durante a coleta das informacoes.

    Returns:
        list[dict[str, Any]]:
            Informacoes das janelas que puderam ser consultadas.

    Raises:
        TypeError:
            Caso as opcoes informadas nao sejam booleanas.
        JanelaError:
            Caso uma janela fique indisponivel e o erro nao seja ignorado.
    """
    # Valida as opcoes booleanas da coleta
    if not isinstance(restaurar_janela_original, bool):
        raise TypeError(
            "O parametro 'restaurar_janela_original' deve ser booleano."
        )

    if not isinstance(ignorar_janelas_indisponiveis, bool):
        raise TypeError(
            "O parametro 'ignorar_janelas_indisponiveis' deve ser booleano."
        )

    # Armazena a janela inicialmente selecionada
    janela_original = obter_janela_atual(driver)
    janelas = listar_janelas(driver)
    informacoes = []

    # Percorre as janelas e coleta os dados individualmente
    for indice, identificador in enumerate(janelas):
        try:
            trocar_para_janela(driver, identificador)
            informacoes.append(
                {
                    "indice": indice,
                    "identificador": identificador,
                    "titulo": driver.title,
                    "url": driver.current_url,
                    "atual": identificador == janela_original,
                }
            )
        except (JanelaNaoEncontradaError, WebDriverException) as error:
            if not ignorar_janelas_indisponiveis:
                raise JanelaError(
                    f"Nao foi possivel consultar a janela '{identificador}'."
                ) from error

    # Restaura o contexto inicial quando solicitado e possivel
    if (
        restaurar_janela_original
        and janela_existe(driver, janela_original)
    ):
        trocar_para_janela(driver, janela_original)

    # Retorna as informacoes coletadas
    return informacoes
