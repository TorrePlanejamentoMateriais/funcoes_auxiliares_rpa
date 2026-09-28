"""
============================================================
Modulo: SELENIUM / navegacao.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para navegacao, consulta de
    informacoes da pagina, historico do navegador e validacao de URLs
    utilizando Selenium.
Funcoes disponiveis:
    - validar_url()
    - normalizar_url()
    - montar_url()
    - obter_dominio_url()
    - obter_parametros_url()
    - adicionar_parametros_url()
    - remover_parametros_url()
    - navegar_para()
    - navegar_para_url_base()
    - voltar_pagina()
    - avancar_pagina()
    - atualizar_pagina()
    - atualizar_pagina_sem_cache()
    - obter_url_atual()
    - obter_titulo_pagina()
    - obter_codigo_fonte()
    - obter_dominio_atual()
    - url_atual_igual()
    - url_atual_contem()
    - titulo_atual_igual()
    - titulo_atual_contem()
    - aguardar_url_igual()
    - aguardar_url_conter()
    - aguardar_url_mudar()
    - aguardar_url_corresponder()
    - aguardar_titulo_igual()
    - aguardar_titulo_conter()
    - aguardar_carregamento_pagina()
    - executar_navegacao()
Dependencias:
    - selenium
    - excecoes
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de navegacao, retorno, avanco e atualizacao.
        - Inclusao de validacao e composicao de URLs.
        - Inclusao de consultas de URL, titulo, dominio e codigo-fonte.
        - Inclusao de esperas para URL, titulo e carregamento da pagina.
        - Inclusao de execucao controlada de fluxos de navegacao.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa Callable para tipar a acao executada apos uma navegacao
from collections.abc import Callable

# Importa funcoes para analisar, montar e modificar URLs
from urllib.parse import (
    parse_qs,
    urlencode,
    urljoin,
    urlparse,
    urlunparse,
)

# Importa Any e TypeVar para tipar retornos genericos
from typing import Any, TypeVar

# Importa excecoes nativas que podem ocorrer durante a navegacao
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    TimeoutException,
    WebDriverException,
)

# Importa o tipo principal do Selenium WebDriver
from selenium.webdriver.remote.webdriver import WebDriver

# Importa as condicoes e o componente de espera explicita
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Importa as excecoes centralizadas da biblioteca
from excecoes import (
    CarregamentoPaginaError,
    DriverInativoError,
    NavegacaoError,
    ParametroInvalidoError,
    TempoEsperaExcedidoError,
    UrlInvalidaError,
)


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------

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

    # Interrompe a execucao quando nao existe identificador de sessao
    if not driver.session_id:
        raise DriverInativoError(
            "O WebDriver informado nao possui uma sessao ativa."
        )

    # Confirma que a janela atual ainda pode ser acessada
    try:
        driver.current_window_handle
    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ) as error:
        raise DriverInativoError(
            "O WebDriver informado nao possui uma sessao ativa."
        ) from error


def _validar_timeout(timeout: int | float) -> None:
    """
    Valida o tempo maximo utilizado nas esperas de navegacao.

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
    # Rejeita booleanos e valores que nao sejam numericos
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
            Nome utilizado nas mensagens de erro.

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

    # Retorna o texto normalizado
    return texto_normalizado


def _validar_booleano(valor: bool, nome_parametro: str) -> None:
    """
    Valida uma opcao booleana recebida por uma funcao.

    Args:
        valor (bool):
            Valor que sera validado.
        nome_parametro (str):
            Nome utilizado na mensagem de erro.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o valor nao seja booleano.
    """
    # Verifica se o valor e exatamente booleano
    if not isinstance(valor, bool):
        raise TypeError(
            f"O parametro '{nome_parametro}' deve ser booleano."
        )


def _comparar_texto(
    valor_atual: str,
    valor_esperado: str,
    correspondencia_exata: bool,
    considerar_maiusculas: bool,
) -> bool:
    """
    Compara dois textos de acordo com as opcoes informadas.

    Args:
        valor_atual (str):
            Texto obtido no navegador.
        valor_esperado (str):
            Texto utilizado como referencia.
        correspondencia_exata (bool):
            Define se os textos devem ser exatamente iguais.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas de minusculas.

    Returns:
        bool:
            True quando os textos correspondem e False caso contrario.
    """
    # Valida as opcoes utilizadas na comparacao
    _validar_booleano(correspondencia_exata, "correspondencia_exata")
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")

    # Normaliza as letras quando a comparacao ignora maiusculas
    atual = valor_atual if considerar_maiusculas else valor_atual.lower()
    esperado = (
        valor_esperado
        if considerar_maiusculas
        else valor_esperado.lower()
    )

    # Executa comparacao exata ou parcial
    if correspondencia_exata:
        return atual == esperado

    return esperado in atual


def _executar_com_tratamento(
    driver: WebDriver,
    operacao: Callable[[], T],
    mensagem_erro: str,
) -> T:
    """
    Executa uma operacao de navegacao e padroniza seus erros.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        operacao (Callable[[], T]):
            Funcao sem argumentos que executa a operacao.
        mensagem_erro (str):
            Mensagem utilizada quando a operacao falhar.

    Returns:
        T:
            Resultado produzido pela operacao.

    Raises:
        ParametroInvalidoError:
            Caso operacao nao seja chamavel.
        NavegacaoError:
            Caso o Selenium nao consiga executar a operacao.
    """
    # Valida a sessao antes de executar a operacao
    _validar_driver(driver)

    # Verifica se a operacao recebida pode ser executada
    if not callable(operacao):
        raise ParametroInvalidoError(
            "O parametro 'operacao' deve ser uma funcao ou objeto chamavel."
        )

    # Valida a mensagem de erro
    mensagem_validada = _validar_texto(mensagem_erro, "mensagem_erro")

    try:
        # Executa a operacao e retorna seu resultado
        return operacao()
    except WebDriverException as error:
        raise NavegacaoError(mensagem_validada) from error


# ----------------------------------------------------------------------------
# FUNCOES DE VALIDACAO E MANIPULACAO DE URL
# ----------------------------------------------------------------------------

def validar_url(
    url: str,
    exigir_protocolo: bool = True,
    protocolos_permitidos: tuple[str, ...] = ("http", "https"),
) -> str:
    """
    Valida uma URL e retorna seu valor normalizado.

    Args:
        url (str):
            URL que sera validada.
        exigir_protocolo (bool):
            Define se a URL deve possuir esquema explicito.
        protocolos_permitidos (tuple[str, ...]):
            Esquemas aceitos, como http e https.

    Returns:
        str:
            URL sem espacos nas extremidades.

    Raises:
        UrlInvalidaError:
            Caso a URL nao possua formato ou protocolo permitido.
        TypeError:
            Caso algum parametro possua tipo invalido.
    """
    # Valida e normaliza o texto da URL
    url_validada = _validar_texto(url, "url")
    _validar_booleano(exigir_protocolo, "exigir_protocolo")

    # Valida a colecao de protocolos permitidos
    if not isinstance(protocolos_permitidos, tuple) or not protocolos_permitidos:
        raise TypeError(
            "O parametro 'protocolos_permitidos' deve ser uma tupla nao vazia."
        )

    if not all(
        isinstance(protocolo, str) and protocolo.strip()
        for protocolo in protocolos_permitidos
    ):
        raise TypeError(
            "Todos os protocolos permitidos devem ser strings nao vazias."
        )

    # Analisa os componentes da URL
    componentes = urlparse(url_validada)
    protocolos = {
        protocolo.strip().lower()
        for protocolo in protocolos_permitidos
    }

    # Exige esquema quando solicitado
    if exigir_protocolo and not componentes.scheme:
        raise UrlInvalidaError(
            f"A URL deve possuir um protocolo explicito: '{url_validada}'."
        )

    # Verifica se o esquema e permitido
    if componentes.scheme and componentes.scheme.lower() not in protocolos:
        raise UrlInvalidaError(
            f"O protocolo '{componentes.scheme}' nao e permitido."
        )

    # Para URLs absolutas, exige um dominio
    if componentes.scheme and not componentes.netloc:
        raise UrlInvalidaError(
            f"A URL nao possui um dominio valido: '{url_validada}'."
        )

    # Para URLs relativas, exige ao menos um caminho
    if not componentes.scheme and not componentes.path:
        raise UrlInvalidaError(
            f"A URL informada e invalida: '{url_validada}'."
        )

    # Retorna a URL validada
    return url_validada


def normalizar_url(
    url: str,
    protocolo_padrao: str = "https",
    remover_barra_final: bool = False,
) -> str:
    """
    Adiciona protocolo quando ausente e normaliza a barra final da URL.

    Args:
        url (str):
            URL absoluta ou dominio que sera normalizado.
        protocolo_padrao (str):
            Protocolo aplicado quando a URL nao possui esquema.
        remover_barra_final (bool):
            Define se a barra final do caminho sera removida.

    Returns:
        str:
            URL absoluta e normalizada.

    Raises:
        UrlInvalidaError:
            Caso a URL final nao seja valida.
    """
    # Valida os textos e a opcao booleana
    url_tratada = _validar_texto(url, "url")
    protocolo = _validar_texto(
        protocolo_padrao,
        "protocolo_padrao",
    ).lower()
    _validar_booleano(remover_barra_final, "remover_barra_final")

    # Adiciona protocolo quando ele nao esta presente
    if not urlparse(url_tratada).scheme:
        url_tratada = f"{protocolo}://{url_tratada}"

    # Valida a URL absoluta resultante
    url_tratada = validar_url(
        url_tratada,
        exigir_protocolo=True,
        protocolos_permitidos=(protocolo, "http", "https"),
    )

    # Analisa os componentes para ajustar o caminho
    componentes = urlparse(url_tratada)
    caminho = componentes.path

    if remover_barra_final and caminho not in {"", "/"}:
        caminho = caminho.rstrip("/")

    # Remonta a URL com o caminho normalizado
    return urlunparse(componentes._replace(path=caminho))


def montar_url(
    url_base: str,
    caminho: str = "",
) -> str:
    """
    Combina uma URL base com um caminho relativo ou absoluto.

    Args:
        url_base (str):
            URL absoluta utilizada como base.
        caminho (str):
            Caminho que sera combinado com a base.

    Returns:
        str:
            URL absoluta resultante da combinacao.

    Raises:
        UrlInvalidaError:
            Caso a URL base ou o resultado sejam invalidos.
    """
    # Valida a URL base e o caminho informado
    base = validar_url(url_base, exigir_protocolo=True)

    if not isinstance(caminho, str):
        raise TypeError(
            "O parametro 'caminho' deve ser uma string."
        )

    # Garante uma barra final para tratar a base como diretorio
    base_para_juncao = base if base.endswith("/") else f"{base}/"

    # Combina a base e o caminho conforme as regras de URL
    resultado = urljoin(base_para_juncao, caminho.strip())

    # Retorna a URL final validada
    return validar_url(resultado, exigir_protocolo=True)


def obter_dominio_url(
    url: str,
    incluir_porta: bool = True,
) -> str:
    """
    Extrai o dominio de uma URL absoluta.

    Args:
        url (str):
            URL cujo dominio sera extraido.
        incluir_porta (bool):
            Define se a porta explicita sera mantida.

    Returns:
        str:
            Dominio com ou sem porta.

    Raises:
        UrlInvalidaError:
            Caso a URL nao possua dominio valido.
    """
    # Valida a URL e a opcao de porta
    url_validada = validar_url(url, exigir_protocolo=True)
    _validar_booleano(incluir_porta, "incluir_porta")
    componentes = urlparse(url_validada)

    # Retorna netloc quando a porta deve ser preservada
    if incluir_porta:
        return componentes.netloc

    # Retorna apenas o hostname sem credenciais ou porta
    if componentes.hostname is None:
        raise UrlInvalidaError(
            f"Nao foi possivel extrair o dominio da URL: '{url_validada}'."
        )

    return componentes.hostname


def obter_parametros_url(url: str) -> dict[str, list[str]]:
    """
    Retorna os parametros de consulta existentes em uma URL.

    Args:
        url (str):
            URL absoluta ou relativa que sera analisada.

    Returns:
        dict[str, list[str]]:
            Parametros e suas listas de valores.

    Raises:
        UrlInvalidaError:
            Caso a URL informada seja invalida.
    """
    # Aceita URL relativa porque a funcao consulta apenas a query string
    url_validada = validar_url(url, exigir_protocolo=False)

    # Converte a query string em um dicionario de listas
    return parse_qs(
        urlparse(url_validada).query,
        keep_blank_values=True,
    )


def adicionar_parametros_url(
    url: str,
    parametros: dict[str, Any],
    sobrescrever: bool = True,
) -> str:
    """
    Adiciona ou atualiza parametros de consulta em uma URL.

    Args:
        url (str):
            URL absoluta ou relativa que sera modificada.
        parametros (dict[str, Any]):
            Parametros adicionados a query string.
        sobrescrever (bool):
            Substitui valores existentes com a mesma chave.

    Returns:
        str:
            URL contendo os parametros resultantes.

    Raises:
        TypeError:
            Caso parametros nao seja um dicionario.
    """
    # Valida a URL, o dicionario e a opcao de sobrescrita
    url_validada = validar_url(url, exigir_protocolo=False)

    if not isinstance(parametros, dict):
        raise TypeError(
            "O parametro 'parametros' deve ser um dicionario."
        )

    _validar_booleano(sobrescrever, "sobrescrever")

    # Recupera os parametros existentes preservando valores repetidos
    componentes = urlparse(url_validada)
    parametros_atuais = parse_qs(
        componentes.query,
        keep_blank_values=True,
    )

    # Adiciona ou combina os novos parametros
    for chave, valor in parametros.items():
        if not isinstance(chave, str) or not chave.strip():
            raise TypeError(
                "Todas as chaves de 'parametros' devem ser strings nao vazias."
            )

        valores = valor if isinstance(valor, (list, tuple)) else [valor]
        valores_texto = [str(item) for item in valores]

        if sobrescrever or chave not in parametros_atuais:
            parametros_atuais[chave] = valores_texto
        else:
            parametros_atuais[chave].extend(valores_texto)

    # Gera a nova query string permitindo parametros com varios valores
    nova_query = urlencode(parametros_atuais, doseq=True)

    # Remonta e retorna a URL
    return urlunparse(componentes._replace(query=nova_query))


def remover_parametros_url(
    url: str,
    parametros: list[str] | tuple[str, ...] | set[str] | None = None,
) -> str:
    """
    Remove parametros especificos ou toda a query string de uma URL.

    Args:
        url (str):
            URL absoluta ou relativa que sera modificada.
        parametros (list | tuple | set | None):
            Chaves removidas. Quando None, remove todos os parametros.

    Returns:
        str:
            URL sem os parametros selecionados.

    Raises:
        TypeError:
            Caso a colecao de parametros seja invalida.
    """
    # Valida a URL permitindo caminhos relativos
    url_validada = validar_url(url, exigir_protocolo=False)
    componentes = urlparse(url_validada)

    # Remove toda a query string quando nenhuma chave e informada
    if parametros is None:
        return urlunparse(componentes._replace(query=""))

    # Valida a colecao de chaves
    if not isinstance(parametros, (list, tuple, set)):
        raise TypeError(
            "O parametro 'parametros' deve ser lista, tupla, set ou None."
        )

    if not all(isinstance(item, str) and item.strip() for item in parametros):
        raise TypeError(
            "Todos os itens de 'parametros' devem ser strings nao vazias."
        )

    # Recupera e filtra os parametros existentes
    parametros_atuais = parse_qs(
        componentes.query,
        keep_blank_values=True,
    )
    chaves_remover = {item.strip() for item in parametros}
    parametros_restantes = {
        chave: valor
        for chave, valor in parametros_atuais.items()
        if chave not in chaves_remover
    }

    # Remonta a URL com os parametros restantes
    return urlunparse(
        componentes._replace(
            query=urlencode(parametros_restantes, doseq=True)
        )
    )


# ----------------------------------------------------------------------------
# FUNCOES DE NAVEGACAO DO NAVEGADOR
# ----------------------------------------------------------------------------

def navegar_para(
    driver: WebDriver,
    url: str,
    aguardar_carregamento: bool = True,
    timeout: int | float = 30,
) -> str:
    """
    Navega para uma URL e opcionalmente aguarda o carregamento completo.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str):
            URL absoluta que sera aberta.
        aguardar_carregamento (bool):
            Define se document.readyState sera aguardado.
        timeout (int | float):
            Tempo maximo para o carregamento da pagina.

    Returns:
        str:
            URL atual apos a navegacao.

    Raises:
        UrlInvalidaError:
            Caso a URL informada seja invalida.
        NavegacaoError:
            Caso o navegador nao consiga acessar a URL.
    """
    # Valida os parametros antes de iniciar a navegacao
    url_validada = validar_url(url, exigir_protocolo=True)
    _validar_booleano(aguardar_carregamento, "aguardar_carregamento")
    _validar_timeout(timeout)

    # Solicita ao WebDriver que abra a URL
    _executar_com_tratamento(
        driver,
        lambda: driver.get(url_validada),
        f"Nao foi possivel navegar para a URL '{url_validada}'.",
    )

    # Aguarda o carregamento quando solicitado
    if aguardar_carregamento:
        aguardar_carregamento_pagina(driver, timeout)

    # Retorna a URL efetivamente aberta
    return obter_url_atual(driver)


def navegar_para_url_base(
    driver: WebDriver,
    url_base: str,
    caminho: str = "",
    aguardar_carregamento: bool = True,
    timeout: int | float = 30,
) -> str:
    """
    Monta uma URL a partir de uma base e navega para o resultado.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url_base (str):
            URL absoluta utilizada como base.
        caminho (str):
            Caminho relativo que sera combinado com a base.
        aguardar_carregamento (bool):
            Define se o carregamento completo sera aguardado.
        timeout (int | float):
            Tempo maximo para o carregamento da pagina.

    Returns:
        str:
            URL atual apos a navegacao.
    """
    # Combina os componentes antes de navegar
    url_final = montar_url(url_base, caminho)

    # Reutiliza a funcao principal de navegacao
    return navegar_para(
        driver=driver,
        url=url_final,
        aguardar_carregamento=aguardar_carregamento,
        timeout=timeout,
    )


def voltar_pagina(
    driver: WebDriver,
    aguardar_carregamento: bool = True,
    timeout: int | float = 30,
) -> str:
    """
    Navega para a pagina anterior do historico do navegador.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        aguardar_carregamento (bool):
            Define se o carregamento da pagina anterior sera aguardado.
        timeout (int | float):
            Tempo maximo para o carregamento.

    Returns:
        str:
            URL atual apos voltar no historico.
    """
    # Valida as opcoes da operacao
    _validar_booleano(aguardar_carregamento, "aguardar_carregamento")
    _validar_timeout(timeout)

    # Executa o comando de retorno do navegador
    _executar_com_tratamento(
        driver,
        driver.back,
        "Nao foi possivel voltar para a pagina anterior.",
    )

    # Aguarda a pagina resultante quando solicitado
    if aguardar_carregamento:
        aguardar_carregamento_pagina(driver, timeout)

    # Retorna a URL apos a movimentacao no historico
    return obter_url_atual(driver)


def avancar_pagina(
    driver: WebDriver,
    aguardar_carregamento: bool = True,
    timeout: int | float = 30,
) -> str:
    """
    Navega para a pagina seguinte do historico do navegador.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        aguardar_carregamento (bool):
            Define se o carregamento da pagina seguinte sera aguardado.
        timeout (int | float):
            Tempo maximo para o carregamento.

    Returns:
        str:
            URL atual apos avancar no historico.
    """
    # Valida as opcoes da operacao
    _validar_booleano(aguardar_carregamento, "aguardar_carregamento")
    _validar_timeout(timeout)

    # Executa o comando de avanco do navegador
    _executar_com_tratamento(
        driver,
        driver.forward,
        "Nao foi possivel avancar para a pagina seguinte.",
    )

    # Aguarda a pagina resultante quando solicitado
    if aguardar_carregamento:
        aguardar_carregamento_pagina(driver, timeout)

    # Retorna a URL apos a movimentacao no historico
    return obter_url_atual(driver)


def atualizar_pagina(
    driver: WebDriver,
    aguardar_carregamento: bool = True,
    timeout: int | float = 30,
) -> str:
    """
    Atualiza a pagina atualmente aberta no navegador.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        aguardar_carregamento (bool):
            Define se o carregamento apos a atualizacao sera aguardado.
        timeout (int | float):
            Tempo maximo para o carregamento.

    Returns:
        str:
            URL atual apos a atualizacao.
    """
    # Valida as opcoes da operacao
    _validar_booleano(aguardar_carregamento, "aguardar_carregamento")
    _validar_timeout(timeout)

    # Solicita a atualizacao pelo WebDriver
    _executar_com_tratamento(
        driver,
        driver.refresh,
        "Nao foi possivel atualizar a pagina atual.",
    )

    # Aguarda a conclusao do carregamento quando solicitado
    if aguardar_carregamento:
        aguardar_carregamento_pagina(driver, timeout)

    # Retorna a URL que permanece aberta
    return obter_url_atual(driver)


def atualizar_pagina_sem_cache(
    driver: WebDriver,
    timeout: int | float = 30,
) -> str:
    """
    Recarrega a pagina ignorando o cache em navegadores Chromium.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        timeout (int | float):
            Tempo maximo para o carregamento apos a atualizacao.

    Returns:
        str:
            URL atual apos o recarregamento.

    Raises:
        NavegacaoError:
            Caso o navegador nao suporte o comando DevTools.
    """
    # Valida a sessao e o tempo da espera
    _validar_driver(driver)
    _validar_timeout(timeout)

    try:
        # Solicita recarregamento sem cache pelo protocolo DevTools
        driver.execute_cdp_cmd(
            "Page.reload",
            {"ignoreCache": True},
        )
    except WebDriverException as error:
        raise NavegacaoError(
            "Nao foi possivel atualizar a pagina sem cache. "
            "Utilize Google Chrome ou Microsoft Edge."
        ) from error

    # Aguarda a conclusao do recarregamento
    aguardar_carregamento_pagina(driver, timeout)

    # Retorna a URL atual
    return obter_url_atual(driver)


# ----------------------------------------------------------------------------
# FUNCOES DE CONSULTA DA PAGINA
# ----------------------------------------------------------------------------

def obter_url_atual(driver: WebDriver) -> str:
    """
    Retorna a URL atualmente aberta na janela selecionada.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            URL atual do navegador.

    Raises:
        NavegacaoError:
            Caso a URL atual nao possa ser consultada.
    """
    # Valida a sessao antes de consultar a URL
    _validar_driver(driver)

    try:
        # Retorna a URL disponibilizada pelo WebDriver
        return driver.current_url
    except WebDriverException as error:
        raise NavegacaoError(
            "Nao foi possivel obter a URL atual."
        ) from error


def obter_titulo_pagina(driver: WebDriver) -> str:
    """
    Retorna o titulo da pagina atualmente aberta.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Titulo atual da pagina.

    Raises:
        NavegacaoError:
            Caso o titulo nao possa ser consultado.
    """
    # Valida a sessao antes de consultar o titulo
    _validar_driver(driver)

    try:
        # Retorna o titulo disponibilizado pelo WebDriver
        return driver.title
    except WebDriverException as error:
        raise NavegacaoError(
            "Nao foi possivel obter o titulo da pagina."
        ) from error


def obter_codigo_fonte(driver: WebDriver) -> str:
    """
    Retorna o codigo-fonte HTML da pagina atual.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Codigo-fonte HTML da pagina atual.

    Raises:
        NavegacaoError:
            Caso o codigo-fonte nao possa ser obtido.
    """
    # Valida a sessao antes de consultar o codigo-fonte
    _validar_driver(driver)

    try:
        # Retorna o codigo-fonte disponibilizado pelo WebDriver
        return driver.page_source
    except WebDriverException as error:
        raise NavegacaoError(
            "Nao foi possivel obter o codigo-fonte da pagina."
        ) from error


def obter_dominio_atual(
    driver: WebDriver,
    incluir_porta: bool = True,
) -> str:
    """
    Retorna o dominio da URL atualmente aberta.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        incluir_porta (bool):
            Define se uma porta explicita sera mantida.

    Returns:
        str:
            Dominio atual com ou sem porta.
    """
    # Recupera a URL atual e reutiliza a extracao de dominio
    return obter_dominio_url(
        obter_url_atual(driver),
        incluir_porta=incluir_porta,
    )


def url_atual_igual(
    driver: WebDriver,
    url: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """
    Verifica se a URL atual e igual ao valor informado.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str):
            URL utilizada na comparacao.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas de minusculas.

    Returns:
        bool:
            True quando as URLs forem iguais e False caso contrario.
    """
    # Valida o texto sem exigir formato absoluto para permitir comparacao livre
    url_validada = _validar_texto(url, "url")

    # Compara a URL atual de forma exata
    return _comparar_texto(
        obter_url_atual(driver),
        url_validada,
        correspondencia_exata=True,
        considerar_maiusculas=considerar_maiusculas,
    )


def url_atual_contem(
    driver: WebDriver,
    texto: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """
    Verifica se a URL atual contem determinado texto.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        texto (str):
            Trecho procurado na URL atual.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas de minusculas.

    Returns:
        bool:
            True quando o trecho estiver presente e False caso contrario.
    """
    # Valida o trecho procurado
    texto_validado = _validar_texto(texto, "texto")

    # Compara a URL atual de forma parcial
    return _comparar_texto(
        obter_url_atual(driver),
        texto_validado,
        correspondencia_exata=False,
        considerar_maiusculas=considerar_maiusculas,
    )


def titulo_atual_igual(
    driver: WebDriver,
    titulo: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """
    Verifica se o titulo atual e igual ao valor informado.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        titulo (str):
            Titulo utilizado na comparacao.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas de minusculas.

    Returns:
        bool:
            True quando os titulos forem iguais e False caso contrario.
    """
    # Valida o titulo esperado
    titulo_validado = _validar_texto(titulo, "titulo")

    # Compara o titulo atual de forma exata
    return _comparar_texto(
        obter_titulo_pagina(driver),
        titulo_validado,
        correspondencia_exata=True,
        considerar_maiusculas=considerar_maiusculas,
    )


def titulo_atual_contem(
    driver: WebDriver,
    texto: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """
    Verifica se o titulo atual contem determinado texto.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        texto (str):
            Trecho procurado no titulo atual.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas de minusculas.

    Returns:
        bool:
            True quando o trecho estiver presente e False caso contrario.
    """
    # Valida o trecho esperado
    texto_validado = _validar_texto(texto, "texto")

    # Compara o titulo atual de forma parcial
    return _comparar_texto(
        obter_titulo_pagina(driver),
        texto_validado,
        correspondencia_exata=False,
        considerar_maiusculas=considerar_maiusculas,
    )


# ----------------------------------------------------------------------------
# FUNCOES DE ESPERA PARA NAVEGACAO
# ----------------------------------------------------------------------------

def _executar_espera_navegacao(
    driver: WebDriver,
    condicao: Callable[[WebDriver], T],
    timeout: int | float,
    mensagem: str,
) -> T:
    """
    Executa uma espera explicita relacionada a navegacao.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        condicao (Callable[[WebDriver], T]):
            Condicao avaliada pelo WebDriverWait.
        timeout (int | float):
            Tempo maximo da espera.
        mensagem (str):
            Mensagem gerada quando o timeout for atingido.

    Returns:
        T:
            Resultado retornado pela condicao atendida.

    Raises:
        TempoEsperaExcedidoError:
            Caso a condicao nao seja atendida no tempo definido.
    """
    # Valida os parametros compartilhados pelas esperas
    _validar_driver(driver)
    _validar_timeout(timeout)

    if not callable(condicao):
        raise ParametroInvalidoError(
            "O parametro 'condicao' deve ser uma funcao ou objeto chamavel."
        )

    mensagem_validada = _validar_texto(mensagem, "mensagem")

    try:
        # Executa a espera explicita ate obter um resultado verdadeiro
        return WebDriverWait(driver, timeout).until(condicao)
    except TimeoutException as error:
        raise TempoEsperaExcedidoError(mensagem_validada) from error


def aguardar_url_igual(
    driver: WebDriver,
    url: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda a URL atual ser exatamente igual ao valor informado.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str):
            URL exata esperada.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando a URL esperada for atingida.
    """
    # Valida a URL absoluta esperada
    url_validada = validar_url(url, exigir_protocolo=True)

    # Executa a condicao oficial de URL exata
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.url_to_be(url_validada),
            timeout,
            f"A URL atual nao ficou igual a '{url_validada}'.",
        )
    )


def aguardar_url_conter(
    driver: WebDriver,
    texto: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda a URL atual conter determinado texto.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        texto (str):
            Trecho esperado na URL.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando o trecho estiver presente na URL.
    """
    # Valida o trecho esperado
    texto_validado = _validar_texto(texto, "texto")

    # Executa a condicao oficial de URL parcial
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.url_contains(texto_validado),
            timeout,
            f"A URL atual nao apresentou '{texto_validado}'.",
        )
    )


def aguardar_url_mudar(
    driver: WebDriver,
    url_anterior: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda a URL atual ficar diferente de uma URL anterior.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url_anterior (str):
            URL existente antes da acao de navegacao.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando a URL for alterada.
    """
    # Valida a URL utilizada como referencia
    url_validada = _validar_texto(url_anterior, "url_anterior")

    # Aguarda a mudanca da URL
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.url_changes(url_validada),
            timeout,
            f"A URL permaneceu igual a '{url_validada}'.",
        )
    )


def aguardar_url_corresponder(
    driver: WebDriver,
    expressao_regular: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda a URL atual corresponder a uma expressao regular.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        expressao_regular (str):
            Expressao regular utilizada na validacao.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando a URL corresponder ao padrao.
    """
    # Valida a expressao regular recebida
    expressao = _validar_texto(
        expressao_regular,
        "expressao_regular",
    )

    # Aguarda a correspondencia da URL
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.url_matches(expressao),
            timeout,
            "A URL nao correspondeu a expressao regular informada.",
        )
    )


def aguardar_titulo_igual(
    driver: WebDriver,
    titulo: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda o titulo da pagina ser exatamente igual ao informado.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        titulo (str):
            Titulo exato esperado.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando o titulo esperado for atingido.
    """
    # Valida o titulo esperado
    titulo_validado = _validar_texto(titulo, "titulo")

    # Aguarda o titulo exato
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.title_is(titulo_validado),
            timeout,
            f"O titulo da pagina nao ficou igual a '{titulo_validado}'.",
        )
    )


def aguardar_titulo_conter(
    driver: WebDriver,
    texto: str,
    timeout: int | float = 20,
) -> bool:
    """
    Aguarda o titulo da pagina conter determinado texto.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        texto (str):
            Trecho esperado no titulo.
        timeout (int | float):
            Tempo maximo da espera.

    Returns:
        bool:
            True quando o trecho estiver presente no titulo.
    """
    # Valida o trecho esperado
    texto_validado = _validar_texto(texto, "texto")

    # Aguarda a presenca do trecho no titulo
    return bool(
        _executar_espera_navegacao(
            driver,
            EC.title_contains(texto_validado),
            timeout,
            f"O titulo da pagina nao apresentou '{texto_validado}'.",
        )
    )


def aguardar_carregamento_pagina(
    driver: WebDriver,
    timeout: int | float = 30,
) -> bool:
    """
    Aguarda document.readyState ficar igual a complete.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        timeout (int | float):
            Tempo maximo para o carregamento da pagina.

    Returns:
        bool:
            True quando o documento concluir o carregamento.

    Raises:
        CarregamentoPaginaError:
            Caso o documento nao fique completo no tempo definido.
    """
    # Define a condicao que consulta o estado do documento
    def pagina_carregada(navegador: WebDriver) -> bool:
        return navegador.execute_script(
            "return document.readyState"
        ) == "complete"

    try:
        # Reutiliza o executor padronizado de espera
        return bool(
            _executar_espera_navegacao(
                driver,
                pagina_carregada,
                timeout,
                "A pagina nao concluiu o carregamento.",
            )
        )
    except TempoEsperaExcedidoError as error:
        # Converte o timeout generico em um erro especifico de carregamento
        raise CarregamentoPaginaError(
            "A pagina nao concluiu o carregamento dentro do tempo limite."
        ) from error


# ----------------------------------------------------------------------------
# FUNCAO DE FLUXO CONTROLADO
# ----------------------------------------------------------------------------

def executar_navegacao(
    driver: WebDriver,
    url: str,
    acao: Callable[[WebDriver], T],
    timeout: int | float = 30,
    retornar_url_anterior: bool = False,
) -> T:
    """
    Navega para uma URL, executa uma acao e opcionalmente retorna a anterior.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        url (str):
            URL absoluta que sera aberta.
        acao (Callable[[WebDriver], T]):
            Funcao executada depois do carregamento da pagina.
        timeout (int | float):
            Tempo maximo para cada carregamento.
        retornar_url_anterior (bool):
            Define se a URL inicial sera restaurada ao final.

    Returns:
        T:
            Resultado retornado pela funcao acao.

    Raises:
        ParametroInvalidoError:
            Caso acao nao seja chamavel.
        NavegacaoError:
            Caso a navegacao ou a restauracao falhe.
    """
    # Valida a funcao e a opcao de restauracao
    if not callable(acao):
        raise ParametroInvalidoError(
            "O parametro 'acao' deve ser uma funcao ou objeto chamavel."
        )

    _validar_booleano(retornar_url_anterior, "retornar_url_anterior")

    # Registra a URL inicial antes da navegacao
    url_anterior = obter_url_atual(driver)

    # Navega para a URL solicitada e aguarda o carregamento
    navegar_para(
        driver=driver,
        url=url,
        aguardar_carregamento=True,
        timeout=timeout,
    )

    try:
        # Executa a acao no contexto da pagina carregada
        return acao(driver)
    finally:
        # Restaura a URL inicial quando solicitado
        if retornar_url_anterior and url_anterior:
            navegar_para(
                driver=driver,
                url=url_anterior,
                aguardar_carregamento=True,
                timeout=timeout,
            )
