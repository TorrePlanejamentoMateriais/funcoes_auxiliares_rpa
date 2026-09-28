"""
============================================================
Modulo: SELENIUM / capturas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para capturas de tela,
    evidencias, codigo-fonte HTML e diagnosticos de paginas web
    utilizando Selenium.
Funcoes disponiveis:
    - preparar_diretorio_capturas()
    - sanitizar_nome_arquivo()
    - gerar_nome_captura()
    - gerar_caminho_captura()
    - capturar_tela()
    - capturar_tela_completa()
    - capturar_elemento()
    - capturar_elemento_web()
    - obter_captura_png()
    - obter_captura_base64()
    - salvar_html_pagina()
    - salvar_logs_navegador()
    - capturar_evidencia_erro()
    - capturar_pacote_evidencias()
    - listar_capturas()
    - obter_captura_mais_recente()
    - remover_capturas_antigas()
Dependencias:
    - selenium
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de capturas da janela e de elementos.
        - Inclusao de captura de pagina completa para navegadores Chromium.
        - Inclusao de evidencias HTML e logs do navegador.
        - Inclusao de pacote de evidencias para tratamento de erros.
        - Inclusao de funcoes para organizacao e limpeza das capturas.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa json para salvar os logs do navegador em formato estruturado
import json

# Importa re para remover caracteres invalidos dos nomes dos arquivos
import re

# Importa datetime para gerar datas e controlar a idade das evidencias
from datetime import datetime, timedelta

# Importa Path para criar e manipular caminhos de arquivos e diretorios
from pathlib import Path

# Importa Any para tipar estruturas genericas retornadas pelo Selenium
from typing import Any

# Importa as excecoes tratadas durante a geracao das evidencias
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    TimeoutException,
    WebDriverException,
)

# Importa os tipos principais utilizados pelo modulo
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

# Importa as condicoes e o componente de espera explicita
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------

# Define o formato padrao de um localizador Selenium
Localizador = tuple[str, str]


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define as extensoes de imagem aceitas pelo modulo
EXTENSOES_IMAGEM = {".png"}

# Define caracteres proibidos ou problematicos em nomes de arquivos
PADRAO_CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ----------------------------------------------------------------------------
# EXCECOES PERSONALIZADAS
# ----------------------------------------------------------------------------

class CapturasError(Exception):
    """Erro base para operacoes relacionadas a capturas e evidencias."""


class DriverInativoError(CapturasError):
    """O WebDriver informado nao possui uma sessao ativa."""


class LocalizadorInvalidoError(CapturasError):
    """O localizador informado nao possui um formato valido."""


class CaminhoCapturaError(CapturasError):
    """O caminho informado para a captura nao e valido."""


class CapturaNaoRealizadaError(CapturasError):
    """O Selenium nao conseguiu gerar ou salvar a captura solicitada."""


class EvidenciaNaoEncontradaError(CapturasError):
    """Nenhuma evidencia correspondente foi encontrada."""


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

    # Tenta acessar a janela atual para confirmar que a sessao responde
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


def _validar_elemento(elemento: WebElement) -> None:
    """
    Valida se o objeto informado e um WebElement.

    Args:
        elemento (WebElement):
            Elemento Selenium que sera validado.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o objeto informado nao seja um WebElement.
    """
    # Verifica se o objeto informado e uma instancia de WebElement
    if not isinstance(elemento, WebElement):
        raise TypeError(
            "O parametro 'elemento' deve ser uma instancia de WebElement."
        )


def _validar_localizador(localizador: Localizador) -> None:
    """
    Valida a estrutura e o conteudo de um localizador Selenium.

    Args:
        localizador (Localizador):
            Tupla contendo a estrategia e o valor da localizacao.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        LocalizadorInvalidoError:
            Caso o localizador nao seja uma tupla com duas strings preenchidas.
    """
    # Verifica se o localizador e uma tupla contendo exatamente dois itens
    if not isinstance(localizador, tuple) or len(localizador) != 2:
        raise LocalizadorInvalidoError(
            "O parametro 'localizador' deve ser uma tupla com dois itens, "
            "por exemplo: (By.ID, 'nome_do_elemento')."
        )

    # Separa a estrategia e o valor para validar cada item
    estrategia, valor = localizador

    # Verifica se os dois itens sao strings preenchidas
    if (
        not isinstance(estrategia, str)
        or not isinstance(valor, str)
        or not estrategia.strip()
        or not valor.strip()
    ):
        raise LocalizadorInvalidoError(
            "A estrategia e o valor do localizador devem ser strings "
            "nao vazias."
        )


def _validar_timeout(timeout: int | float) -> None:
    """
    Valida o tempo maximo utilizado para localizar um elemento.

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
    # Rejeita booleanos porque bool tambem e considerado int no Python
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)):
        raise TypeError(
            "O parametro 'timeout' deve ser numerico."
        )

    # Verifica se o timeout e maior que zero
    if timeout <= 0:
        raise ValueError(
            "O parametro 'timeout' deve ser maior que zero."
        )


def _validar_extensao_png(caminho: Path) -> None:
    """
    Valida se o caminho da captura utiliza a extensao PNG.

    Args:
        caminho (Path):
            Caminho que sera validado.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        CaminhoCapturaError:
            Caso o caminho nao utilize a extensao .png.
    """
    # Compara a extensao em letras minusculas com as extensoes aceitas
    if caminho.suffix.lower() not in EXTENSOES_IMAGEM:
        raise CaminhoCapturaError(
            "As capturas de tela devem utilizar a extensao '.png'."
        )


def _preparar_caminho_arquivo(
    caminho: str | Path,
    extensao_obrigatoria: str | None = None,
) -> Path:
    """
    Normaliza o caminho e cria o diretorio pai do arquivo.

    Args:
        caminho (str | Path):
            Caminho do arquivo que sera preparado.
        extensao_obrigatoria (str | None):
            Extensao que deve ser aplicada quando o caminho nao possuir uma.

    Returns:
        Path:
            Caminho absoluto e preparado para gravacao.

    Raises:
        TypeError:
            Caso o caminho nao seja uma string ou Path.
        CaminhoCapturaError:
            Caso o caminho represente um diretorio existente.
    """
    # Verifica o tipo do caminho informado
    if not isinstance(caminho, (str, Path)):
        raise TypeError(
            "O parametro 'caminho' deve ser uma string ou Path."
        )

    # Converte o caminho informado para um caminho absoluto
    caminho_preparado = Path(caminho).expanduser().resolve()

    # Interrompe a execucao quando o caminho aponta para um diretorio
    if caminho_preparado.exists() and caminho_preparado.is_dir():
        raise CaminhoCapturaError(
            f"O caminho informado representa um diretorio: {caminho_preparado}"
        )

    # Adiciona a extensao quando ela for obrigatoria e estiver ausente
    if extensao_obrigatoria and not caminho_preparado.suffix:
        caminho_preparado = caminho_preparado.with_suffix(
            extensao_obrigatoria
        )

    # Cria o diretorio pai quando ele ainda nao existe
    caminho_preparado.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Retorna o caminho preparado
    return caminho_preparado


def _validar_dias(dias: int) -> None:
    """
    Valida uma quantidade de dias utilizada na limpeza de capturas.

    Args:
        dias (int):
            Quantidade de dias que sera validada.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso dias nao seja um numero inteiro.
        ValueError:
            Caso dias seja menor que zero.
    """
    # Rejeita booleanos e valores que nao sejam inteiros
    if isinstance(dias, bool) or not isinstance(dias, int):
        raise TypeError(
            "O parametro 'dias' deve ser um numero inteiro."
        )

    # Permite zero para remover evidencias anteriores ao momento atual
    if dias < 0:
        raise ValueError(
            "O parametro 'dias' nao pode ser negativo."
        )


# ----------------------------------------------------------------------------
# FUNCOES DE NOMES E DIRETORIOS
# ----------------------------------------------------------------------------

def preparar_diretorio_capturas(
    diretorio: str | Path = "capturas",
) -> Path:
    """
    Cria e retorna o diretorio utilizado para armazenar as capturas.

    Args:
        diretorio (str | Path):
            Diretorio onde as evidencias serao armazenadas.

    Returns:
        Path:
            Caminho absoluto do diretorio criado ou localizado.

    Raises:
        TypeError:
            Caso o diretorio nao seja uma string ou Path.
        CaminhoCapturaError:
            Caso o caminho exista e nao seja um diretorio.
        OSError:
            Caso o sistema operacional nao consiga criar o diretorio.
    """
    # Verifica o tipo do diretorio informado
    if not isinstance(diretorio, (str, Path)):
        raise TypeError(
            "O parametro 'diretorio' deve ser uma string ou Path."
        )

    # Converte o diretorio para um caminho absoluto
    caminho_diretorio = Path(diretorio).expanduser().resolve()

    # Interrompe a execucao quando existe um arquivo no caminho desejado
    if caminho_diretorio.exists() and not caminho_diretorio.is_dir():
        raise CaminhoCapturaError(
            f"O caminho informado nao e um diretorio: {caminho_diretorio}"
        )

    # Cria o diretorio e seus pais quando necessario
    caminho_diretorio.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Retorna o diretorio preparado
    return caminho_diretorio


def sanitizar_nome_arquivo(
    nome: str,
    substituto: str = "_",
    tamanho_maximo: int = 120,
) -> str:
    """
    Remove caracteres invalidos e normaliza um nome de arquivo.

    Args:
        nome (str):
            Nome original que sera sanitizado.
        substituto (str):
            Caractere utilizado no lugar dos caracteres invalidos.
        tamanho_maximo (int):
            Quantidade maxima de caracteres do nome retornado.

    Returns:
        str:
            Nome seguro para utilizacao no sistema de arquivos.

    Raises:
        TypeError:
            Caso os parametros possuam tipos invalidos.
        ValueError:
            Caso nome, substituto ou tamanho_maximo sejam invalidos.
    """
    # Verifica se o nome informado e uma string
    if not isinstance(nome, str):
        raise TypeError(
            "O parametro 'nome' deve ser uma string."
        )

    # Verifica se o substituto e uma string
    if not isinstance(substituto, str):
        raise TypeError(
            "O parametro 'substituto' deve ser uma string."
        )

    # Verifica se o tamanho maximo e um inteiro positivo
    if (
        isinstance(tamanho_maximo, bool)
        or not isinstance(tamanho_maximo, int)
    ):
        raise TypeError(
            "O parametro 'tamanho_maximo' deve ser um numero inteiro."
        )

    if tamanho_maximo <= 0:
        raise ValueError(
            "O parametro 'tamanho_maximo' deve ser maior que zero."
        )

    # Remove espacos nas extremidades do nome
    nome_tratado = nome.strip()

    # Interrompe a execucao quando o nome original esta vazio
    if not nome_tratado:
        raise ValueError(
            "O parametro 'nome' nao pode estar vazio."
        )

    # Substitui os caracteres proibidos no sistema de arquivos
    nome_tratado = PADRAO_CARACTERES_INVALIDOS.sub(
        substituto,
        nome_tratado,
    )

    # Substitui sequencias de espacos por um unico sublinhado
    nome_tratado = re.sub(r"\s+", "_", nome_tratado)

    # Remove pontos e espacos que podem causar problemas no Windows
    nome_tratado = nome_tratado.strip(" ._")

    # Interrompe a execucao quando nao resta nenhum caractere valido
    if not nome_tratado:
        raise ValueError(
            "O nome informado nao possui caracteres validos."
        )

    # Limita o tamanho final do nome
    return nome_tratado[:tamanho_maximo]


def gerar_nome_captura(
    descricao: str = "captura",
    prefixo: str | None = None,
    data_hora: datetime | None = None,
    extensao: str = ".png",
) -> str:
    """
    Gera um nome padronizado para uma captura ou evidencia.

    Args:
        descricao (str):
            Descricao utilizada para identificar a evidencia.
        prefixo (str | None):
            Texto opcional adicionado antes da data e hora.
        data_hora (datetime | None):
            Data e hora utilizadas no nome. Quando None, utiliza o momento atual.
        extensao (str):
            Extensao adicionada ao nome gerado.

    Returns:
        str:
            Nome padronizado no formato prefixo_data_descricao.extensao.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        ValueError:
            Caso a extensao esteja vazia.
    """
    # Valida a data e hora quando ela for informada manualmente
    if data_hora is not None and not isinstance(data_hora, datetime):
        raise TypeError(
            "O parametro 'data_hora' deve ser datetime ou None."
        )

    # Verifica se a extensao informada e uma string
    if not isinstance(extensao, str):
        raise TypeError(
            "O parametro 'extensao' deve ser uma string."
        )

    # Remove espacos e garante que a extensao comece com ponto
    extensao_tratada = extensao.strip().lower()

    if not extensao_tratada:
        raise ValueError(
            "O parametro 'extensao' nao pode estar vazio."
        )

    if not extensao_tratada.startswith("."):
        extensao_tratada = f".{extensao_tratada}"

    # Sanitiza a descricao principal da evidencia
    descricao_tratada = sanitizar_nome_arquivo(descricao)

    # Utiliza a data informada ou o momento atual
    momento = data_hora or datetime.now()
    data_formatada = momento.strftime("%Y-%m-%d_%H-%M-%S-%f")

    # Monta as partes do nome antes de aplicar a extensao
    partes = []

    if prefixo is not None:
        partes.append(sanitizar_nome_arquivo(prefixo))

    partes.extend([data_formatada, descricao_tratada])

    # Retorna o nome completo da evidencia
    return "_".join(partes) + extensao_tratada


def gerar_caminho_captura(
    diretorio: str | Path = "capturas",
    descricao: str = "captura",
    prefixo: str | None = None,
    data_hora: datetime | None = None,
) -> Path:
    """
    Gera o caminho completo para uma nova captura PNG.

    Args:
        diretorio (str | Path):
            Diretorio onde a captura sera armazenada.
        descricao (str):
            Descricao utilizada no nome da captura.
        prefixo (str | None):
            Prefixo opcional adicionado ao nome.
        data_hora (datetime | None):
            Data e hora utilizadas na geracao do nome.

    Returns:
        Path:
            Caminho absoluto para a nova captura PNG.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        CaminhoCapturaError:
            Caso o diretorio informado seja invalido.
    """
    # Prepara o diretorio onde a evidencia sera armazenada
    diretorio_preparado = preparar_diretorio_capturas(diretorio)

    # Gera um nome padronizado para a captura
    nome_arquivo = gerar_nome_captura(
        descricao=descricao,
        prefixo=prefixo,
        data_hora=data_hora,
        extensao=".png",
    )

    # Retorna o caminho completo
    return diretorio_preparado / nome_arquivo


# ----------------------------------------------------------------------------
# FUNCOES DE CAPTURA DE TELA
# ----------------------------------------------------------------------------

def capturar_tela(
    driver: WebDriver,
    caminho: str | Path,
) -> Path:
    """
    Captura a area visivel da janela atual do navegador.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        caminho (str | Path):
            Caminho onde a captura PNG sera salva.

    Returns:
        Path:
            Caminho absoluto da captura salva.

    Raises:
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        CaminhoCapturaError:
            Caso o caminho ou a extensao sejam invalidos.
        CapturaNaoRealizadaError:
            Caso o Selenium nao consiga salvar a captura.
    """
    # Valida a sessao antes de solicitar a captura
    _validar_driver(driver)

    # Prepara o caminho e garante a extensao PNG
    caminho_preparado = _preparar_caminho_arquivo(
        caminho,
        extensao_obrigatoria=".png",
    )
    _validar_extensao_png(caminho_preparado)

    try:
        # Solicita ao Selenium a captura da janela visivel
        captura_realizada = driver.save_screenshot(
            str(caminho_preparado)
        )
    except WebDriverException as error:
        raise CapturaNaoRealizadaError(
            "O Selenium nao conseguiu capturar a tela atual."
        ) from error

    # O Selenium retorna False quando nao consegue gravar o arquivo
    if not captura_realizada or not caminho_preparado.is_file():
        raise CapturaNaoRealizadaError(
            f"A captura nao foi salva em: {caminho_preparado}"
        )

    # Retorna o caminho da evidencia criada
    return caminho_preparado


def capturar_tela_completa(
    driver: WebDriver,
    caminho: str | Path,
    restaurar_metricas: bool = True,
) -> Path:
    """
    Captura a pagina completa utilizando o protocolo DevTools do Chromium.

    A funcao e indicada para Chrome e Edge. Quando o comando DevTools nao
    estiver disponivel, uma CapturaNaoRealizadaError sera gerada.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        caminho (str | Path):
            Caminho onde a captura PNG sera salva.
        restaurar_metricas (bool):
            Define se as metricas emuladas serao removidas ao final.

    Returns:
        Path:
            Caminho absoluto da captura completa salva.

    Raises:
        TypeError:
            Caso restaurar_metricas nao seja booleano.
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        CapturaNaoRealizadaError:
            Caso o navegador nao ofereca suporte ao comando DevTools.
    """
    # Valida a opcao de restauracao das metricas
    if not isinstance(restaurar_metricas, bool):
        raise TypeError(
            "O parametro 'restaurar_metricas' deve ser booleano."
        )

    # Valida o driver e prepara o caminho de destino
    _validar_driver(driver)
    caminho_preparado = _preparar_caminho_arquivo(
        caminho,
        extensao_obrigatoria=".png",
    )
    _validar_extensao_png(caminho_preparado)

    try:
        # Consulta as dimensoes completas do conteudo da pagina
        metricas = driver.execute_cdp_cmd(
            "Page.getLayoutMetrics",
            {},
        )
        tamanho = metricas.get("contentSize", {})
        largura = max(1, int(tamanho.get("width", 1)))
        altura = max(1, int(tamanho.get("height", 1)))

        # Configura a area emulada para abranger toda a pagina
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {
                "mobile": False,
                "width": largura,
                "height": altura,
                "deviceScaleFactor": 1,
            },
        )

        # Solicita a captura em PNG e recebe o conteudo em base64
        resultado = driver.execute_cdp_cmd(
            "Page.captureScreenshot",
            {
                "format": "png",
                "captureBeyondViewport": True,
                "fromSurface": True,
            },
        )

        # Recupera o conteudo base64 retornado pelo navegador
        conteudo_base64 = resultado.get("data")

        if not conteudo_base64:
            raise CapturaNaoRealizadaError(
                "O navegador nao retornou os dados da captura completa."
            )

        # Importa base64 localmente porque ele e usado somente nesta funcao
        import base64

        # Converte o conteudo base64 para bytes e salva o arquivo
        caminho_preparado.write_bytes(
            base64.b64decode(conteudo_base64)
        )

    except CapturaNaoRealizadaError:
        raise
    except (WebDriverException, ValueError, TypeError, KeyError) as error:
        raise CapturaNaoRealizadaError(
            "Nao foi possivel capturar a pagina completa. "
            "Utilize Chrome ou Edge com suporte ao protocolo DevTools."
        ) from error
    finally:
        # Remove a emulacao de metricas quando solicitado
        if restaurar_metricas:
            try:
                driver.execute_cdp_cmd(
                    "Emulation.clearDeviceMetricsOverride",
                    {},
                )
            except WebDriverException:
                pass

    # Confirma que o arquivo foi efetivamente criado
    if not caminho_preparado.is_file():
        raise CapturaNaoRealizadaError(
            f"A captura completa nao foi salva em: {caminho_preparado}"
        )

    # Retorna o caminho da captura completa
    return caminho_preparado


def capturar_elemento_web(
    elemento: WebElement,
    caminho: str | Path,
) -> Path:
    """
    Captura um WebElement que ja foi localizado anteriormente.

    Args:
        elemento (WebElement):
            Elemento Selenium que sera capturado.
        caminho (str | Path):
            Caminho onde a captura PNG sera salva.

    Returns:
        Path:
            Caminho absoluto da captura do elemento.

    Raises:
        TypeError:
            Caso elemento nao seja um WebElement.
        CaminhoCapturaError:
            Caso o caminho da captura seja invalido.
        CapturaNaoRealizadaError:
            Caso o Selenium nao consiga salvar a captura.
    """
    # Valida o objeto recebido antes de realizar a captura
    _validar_elemento(elemento)

    # Prepara o caminho e valida sua extensao
    caminho_preparado = _preparar_caminho_arquivo(
        caminho,
        extensao_obrigatoria=".png",
    )
    _validar_extensao_png(caminho_preparado)

    try:
        # Solicita ao proprio WebElement a captura de sua area
        captura_realizada = elemento.screenshot(
            str(caminho_preparado)
        )
    except WebDriverException as error:
        raise CapturaNaoRealizadaError(
            "O Selenium nao conseguiu capturar o elemento."
        ) from error

    # Confirma o retorno do Selenium e a existencia do arquivo
    if not captura_realizada or not caminho_preparado.is_file():
        raise CapturaNaoRealizadaError(
            f"A captura do elemento nao foi salva em: {caminho_preparado}"
        )

    # Retorna o caminho da captura criada
    return caminho_preparado


def capturar_elemento(
    driver: WebDriver,
    localizador: Localizador,
    caminho: str | Path,
    timeout: int | float = 20,
    rolar_ate_elemento: bool = True,
) -> Path:
    """
    Localiza um elemento visivel e salva uma captura de sua area.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        localizador (Localizador):
            Tupla contendo a estrategia e o valor da localizacao.
        caminho (str | Path):
            Caminho onde a captura PNG sera salva.
        timeout (int | float):
            Tempo maximo para o elemento ficar visivel.
        rolar_ate_elemento (bool):
            Define se a pagina sera rolada ate o elemento antes da captura.

    Returns:
        Path:
            Caminho absoluto da captura do elemento.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        LocalizadorInvalidoError:
            Caso o localizador seja invalido.
        CapturaNaoRealizadaError:
            Caso o elemento nao fique visivel ou a captura falhe.
    """
    # Valida os parametros utilizados na localizacao
    _validar_driver(driver)
    _validar_localizador(localizador)
    _validar_timeout(timeout)

    # Verifica se a opcao de rolagem e booleana
    if not isinstance(rolar_ate_elemento, bool):
        raise TypeError(
            "O parametro 'rolar_ate_elemento' deve ser booleano."
        )

    try:
        # Aguarda o elemento ficar visivel antes de captura-lo
        elemento = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located(localizador)
        )

        # Centraliza o elemento na area visivel quando solicitado
        if rolar_ate_elemento:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                elemento,
            )

        # Reutiliza a funcao responsavel pela captura de WebElement
        return capturar_elemento_web(
            elemento=elemento,
            caminho=caminho,
        )

    except TimeoutException as error:
        raise CapturaNaoRealizadaError(
            f"O elemento {localizador} nao ficou visivel dentro de "
            f"{timeout} segundo(s)."
        ) from error
    except WebDriverException as error:
        raise CapturaNaoRealizadaError(
            f"Nao foi possivel capturar o elemento {localizador}."
        ) from error


def obter_captura_png(driver: WebDriver) -> bytes:
    """
    Retorna a captura da area visivel em formato de bytes PNG.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        bytes:
            Conteudo binario da captura PNG.

    Raises:
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        CapturaNaoRealizadaError:
            Caso o navegador nao consiga gerar a captura.
    """
    # Valida a sessao antes de solicitar os bytes
    _validar_driver(driver)

    try:
        # Solicita ao Selenium os bytes da captura atual
        conteudo = driver.get_screenshot_as_png()
    except WebDriverException as error:
        raise CapturaNaoRealizadaError(
            "O Selenium nao conseguiu gerar a captura em bytes."
        ) from error

    # Verifica se o retorno possui conteudo binario
    if not isinstance(conteudo, bytes) or not conteudo:
        raise CapturaNaoRealizadaError(
            "O navegador retornou uma captura PNG vazia ou invalida."
        )

    # Retorna os bytes da captura
    return conteudo


def obter_captura_base64(driver: WebDriver) -> str:
    """
    Retorna a captura da area visivel como uma string base64.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.

    Returns:
        str:
            Conteudo base64 da captura atual.

    Raises:
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        CapturaNaoRealizadaError:
            Caso o navegador nao consiga gerar a captura.
    """
    # Valida a sessao antes de solicitar a captura base64
    _validar_driver(driver)

    try:
        # Solicita ao Selenium a captura codificada em base64
        conteudo = driver.get_screenshot_as_base64()
    except WebDriverException as error:
        raise CapturaNaoRealizadaError(
            "O Selenium nao conseguiu gerar a captura em base64."
        ) from error

    # Verifica se o retorno e uma string preenchida
    if not isinstance(conteudo, str) or not conteudo.strip():
        raise CapturaNaoRealizadaError(
            "O navegador retornou uma captura base64 vazia ou invalida."
        )

    # Retorna a string base64 sem espacos nas extremidades
    return conteudo.strip()


# ----------------------------------------------------------------------------
# FUNCOES DE EVIDENCIAS COMPLEMENTARES
# ----------------------------------------------------------------------------

def salvar_html_pagina(
    driver: WebDriver,
    caminho: str | Path,
    encoding: str = "utf-8",
) -> Path:
    """
    Salva o codigo-fonte HTML atual para analise e diagnostico.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        caminho (str | Path):
            Caminho onde o arquivo HTML sera salvo.
        encoding (str):
            Codificacao utilizada na gravacao do arquivo.

    Returns:
        Path:
            Caminho absoluto do arquivo HTML salvo.

    Raises:
        TypeError:
            Caso encoding nao seja uma string.
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        CapturaNaoRealizadaError:
            Caso nao seja possivel obter ou salvar o HTML.
    """
    # Valida a sessao ativa
    _validar_driver(driver)

    # Verifica se a codificacao foi informada corretamente
    if not isinstance(encoding, str):
        raise TypeError(
            "O parametro 'encoding' deve ser uma string."
        )

    if not encoding.strip():
        raise ValueError(
            "O parametro 'encoding' nao pode estar vazio."
        )

    # Prepara o caminho e adiciona a extensao HTML quando necessario
    caminho_preparado = _preparar_caminho_arquivo(
        caminho,
        extensao_obrigatoria=".html",
    )

    # Impede a gravacao acidental com uma extensao diferente
    if caminho_preparado.suffix.lower() not in {".html", ".htm"}:
        raise CaminhoCapturaError(
            "O codigo-fonte deve ser salvo com extensao '.html' ou '.htm'."
        )

    try:
        # Recupera o codigo-fonte da pagina atual
        html = driver.page_source

        # Salva o codigo-fonte utilizando a codificacao informada
        caminho_preparado.write_text(
            html,
            encoding=encoding.strip(),
        )
    except (WebDriverException, OSError, LookupError) as error:
        raise CapturaNaoRealizadaError(
            "Nao foi possivel salvar o codigo-fonte da pagina."
        ) from error

    # Retorna o caminho do HTML salvo
    return caminho_preparado


def salvar_logs_navegador(
    driver: WebDriver,
    caminho: str | Path,
    tipo_log: str = "browser",
    encoding: str = "utf-8",
) -> Path:
    """
    Salva os logs disponibilizados pelo navegador em formato JSON.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        caminho (str | Path):
            Caminho onde o arquivo JSON sera salvo.
        tipo_log (str):
            Tipo de log solicitado ao Selenium, como browser ou performance.
        encoding (str):
            Codificacao utilizada na gravacao do arquivo.

    Returns:
        Path:
            Caminho absoluto do arquivo JSON salvo.

    Raises:
        TypeError:
            Caso tipo_log ou encoding possuam tipos invalidos.
        CapturaNaoRealizadaError:
            Caso os logs nao estejam disponiveis ou nao possam ser salvos.
    """
    # Valida a sessao ativa antes de consultar os logs
    _validar_driver(driver)

    # Verifica os parametros textuais obrigatorios
    if not isinstance(tipo_log, str) or not tipo_log.strip():
        raise TypeError(
            "O parametro 'tipo_log' deve ser uma string nao vazia."
        )

    if not isinstance(encoding, str) or not encoding.strip():
        raise TypeError(
            "O parametro 'encoding' deve ser uma string nao vazia."
        )

    # Prepara o caminho de destino em formato JSON
    caminho_preparado = _preparar_caminho_arquivo(
        caminho,
        extensao_obrigatoria=".json",
    )

    if caminho_preparado.suffix.lower() != ".json":
        raise CaminhoCapturaError(
            "Os logs do navegador devem ser salvos com extensao '.json'."
        )

    try:
        # Recupera os registros do tipo solicitado
        logs = driver.get_log(tipo_log.strip())

        # Salva os logs preservando caracteres acentuados
        caminho_preparado.write_text(
            json.dumps(
                logs,
                ensure_ascii=False,
                indent=4,
                default=str,
            ),
            encoding=encoding.strip(),
        )
    except (WebDriverException, OSError, LookupError, TypeError) as error:
        raise CapturaNaoRealizadaError(
            f"Nao foi possivel salvar os logs do tipo '{tipo_log}'."
        ) from error

    # Retorna o caminho do arquivo de logs
    return caminho_preparado


def capturar_evidencia_erro(
    driver: WebDriver,
    diretorio: str | Path = "capturas",
    etapa: str = "erro",
    incluir_html: bool = True,
) -> dict[str, Path | None]:
    """
    Captura uma evidencia padronizada quando ocorre uma falha na automacao.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        diretorio (str | Path):
            Diretorio onde as evidencias serao armazenadas.
        etapa (str):
            Nome da etapa ou processo que apresentou a falha.
        incluir_html (bool):
            Define se o codigo-fonte HTML tambem sera salvo.

    Returns:
        dict[str, Path | None]:
            Caminhos da imagem e do HTML gerados.

    Raises:
        TypeError:
            Caso incluir_html nao seja booleano.
        CapturaNaoRealizadaError:
            Caso a captura principal nao possa ser criada.
    """
    # Verifica a opcao de inclusao do HTML
    if not isinstance(incluir_html, bool):
        raise TypeError(
            "O parametro 'incluir_html' deve ser booleano."
        )

    # Prepara o diretorio base das evidencias
    diretorio_preparado = preparar_diretorio_capturas(diretorio)

    # Gera uma data unica para manter os arquivos relacionados agrupados
    momento = datetime.now()
    descricao = sanitizar_nome_arquivo(etapa)

    # Gera e salva a captura principal
    caminho_imagem = diretorio_preparado / gerar_nome_captura(
        descricao=descricao,
        prefixo="erro",
        data_hora=momento,
        extensao=".png",
    )
    capturar_tela(driver, caminho_imagem)

    # Inicializa o caminho HTML como None
    caminho_html = None

    # Salva o HTML quando solicitado
    if incluir_html:
        caminho_html = diretorio_preparado / gerar_nome_captura(
            descricao=descricao,
            prefixo="erro",
            data_hora=momento,
            extensao=".html",
        )
        salvar_html_pagina(driver, caminho_html)

    # Retorna os caminhos gerados
    return {
        "imagem": caminho_imagem,
        "html": caminho_html,
    }


def capturar_pacote_evidencias(
    driver: WebDriver,
    diretorio: str | Path = "capturas",
    descricao: str = "evidencia",
    incluir_html: bool = True,
    incluir_logs: bool = False,
    tipo_log: str = "browser",
) -> dict[str, Path | None]:
    """
    Gera um pacote com captura, HTML e logs opcionais da pagina atual.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        diretorio (str | Path):
            Diretorio onde o pacote sera armazenado.
        descricao (str):
            Descricao utilizada nos nomes das evidencias.
        incluir_html (bool):
            Define se o codigo-fonte HTML sera incluido.
        incluir_logs (bool):
            Define se os logs do navegador serao incluidos.
        tipo_log (str):
            Tipo de log solicitado quando incluir_logs for True.

    Returns:
        dict[str, Path | None]:
            Caminhos da imagem, HTML e logs gerados.

    Raises:
        TypeError:
            Caso as opcoes booleanas possuam tipos invalidos.
        CapturaNaoRealizadaError:
            Caso uma das evidencias solicitadas nao possa ser criada.
    """
    # Valida as opcoes booleanas do pacote
    if not isinstance(incluir_html, bool):
        raise TypeError(
            "O parametro 'incluir_html' deve ser booleano."
        )

    if not isinstance(incluir_logs, bool):
        raise TypeError(
            "O parametro 'incluir_logs' deve ser booleano."
        )

    # Prepara o diretorio e os dados compartilhados pelos arquivos
    diretorio_preparado = preparar_diretorio_capturas(diretorio)
    momento = datetime.now()
    descricao_tratada = sanitizar_nome_arquivo(descricao)

    # Gera a captura obrigatoria do pacote
    caminho_imagem = diretorio_preparado / gerar_nome_captura(
        descricao=descricao_tratada,
        data_hora=momento,
        extensao=".png",
    )
    capturar_tela(driver, caminho_imagem)

    # Inicializa evidencias opcionais
    caminho_html = None
    caminho_logs = None

    # Gera o arquivo HTML quando solicitado
    if incluir_html:
        caminho_html = diretorio_preparado / gerar_nome_captura(
            descricao=descricao_tratada,
            data_hora=momento,
            extensao=".html",
        )
        salvar_html_pagina(driver, caminho_html)

    # Gera o arquivo de logs quando solicitado
    if incluir_logs:
        caminho_logs = diretorio_preparado / gerar_nome_captura(
            descricao=descricao_tratada,
            data_hora=momento,
            extensao=".json",
        )
        salvar_logs_navegador(
            driver=driver,
            caminho=caminho_logs,
            tipo_log=tipo_log,
        )

    # Retorna o pacote de caminhos gerados
    return {
        "imagem": caminho_imagem,
        "html": caminho_html,
        "logs": caminho_logs,
    }


# ----------------------------------------------------------------------------
# FUNCOES DE CONSULTA E LIMPEZA
# ----------------------------------------------------------------------------

def listar_capturas(
    diretorio: str | Path = "capturas",
    extensoes: tuple[str, ...] = (".png",),
    recursivo: bool = False,
) -> list[Path]:
    """
    Lista e ordena as capturas existentes em um diretorio.

    Args:
        diretorio (str | Path):
            Diretorio consultado.
        extensoes (tuple[str, ...]):
            Extensoes que serao consideradas na consulta.
        recursivo (bool):
            Define se subdiretorios tambem serao pesquisados.

    Returns:
        list[Path]:
            Lista de arquivos ordenada pela data de modificacao mais recente.

    Raises:
        TypeError:
            Caso extensoes ou recursivo possuam tipos invalidos.
        CaminhoCapturaError:
            Caso o diretorio informado nao exista.
    """
    # Verifica se a pesquisa devera ser recursiva
    if not isinstance(recursivo, bool):
        raise TypeError(
            "O parametro 'recursivo' deve ser booleano."
        )

    # Valida a estrutura da lista de extensoes
    if not isinstance(extensoes, tuple) or not extensoes:
        raise TypeError(
            "O parametro 'extensoes' deve ser uma tupla nao vazia."
        )

    if not all(
        isinstance(extensao, str) and extensao.strip()
        for extensao in extensoes
    ):
        raise TypeError(
            "Todos os itens de 'extensoes' devem ser strings nao vazias."
        )

    # Valida o tipo do diretorio informado
    if not isinstance(diretorio, (str, Path)):
        raise TypeError(
            "O parametro 'diretorio' deve ser uma string ou Path."
        )

    # Converte o diretorio para caminho absoluto
    caminho_diretorio = Path(diretorio).expanduser().resolve()

    # Interrompe a execucao quando o diretorio nao existe
    if not caminho_diretorio.is_dir():
        raise CaminhoCapturaError(
            f"O diretorio de capturas nao foi encontrado: {caminho_diretorio}"
        )

    # Normaliza as extensoes para comparacao
    extensoes_normalizadas = {
        extensao.lower()
        if extensao.startswith(".")
        else f".{extensao.lower()}"
        for extensao in extensoes
    }

    # Seleciona o metodo de pesquisa conforme a opcao recursiva
    arquivos_candidatos = (
        caminho_diretorio.rglob("*")
        if recursivo
        else caminho_diretorio.glob("*")
    )

    # Filtra somente arquivos com as extensoes solicitadas
    arquivos = [
        arquivo
        for arquivo in arquivos_candidatos
        if arquivo.is_file()
        and arquivo.suffix.lower() in extensoes_normalizadas
    ]

    # Ordena do arquivo mais recente para o mais antigo
    return sorted(
        arquivos,
        key=lambda arquivo: arquivo.stat().st_mtime,
        reverse=True,
    )


def obter_captura_mais_recente(
    diretorio: str | Path = "capturas",
    extensoes: tuple[str, ...] = (".png",),
    recursivo: bool = False,
) -> Path:
    """
    Retorna a captura modificada mais recentemente.

    Args:
        diretorio (str | Path):
            Diretorio consultado.
        extensoes (tuple[str, ...]):
            Extensoes consideradas na consulta.
        recursivo (bool):
            Define se subdiretorios serao pesquisados.

    Returns:
        Path:
            Caminho da evidencia mais recente.

    Raises:
        EvidenciaNaoEncontradaError:
            Caso nenhuma evidencia correspondente seja encontrada.
        CaminhoCapturaError:
            Caso o diretorio informado nao exista.
    """
    # Reutiliza a listagem ordenada das capturas
    capturas = listar_capturas(
        diretorio=diretorio,
        extensoes=extensoes,
        recursivo=recursivo,
    )

    # Interrompe a execucao quando nao existem evidencias
    if not capturas:
        raise EvidenciaNaoEncontradaError(
            "Nenhuma captura correspondente foi encontrada."
        )

    # Retorna o primeiro item porque a lista esta em ordem decrescente
    return capturas[0]


def remover_capturas_antigas(
    diretorio: str | Path = "capturas",
    dias: int = 30,
    extensoes: tuple[str, ...] = (".png", ".html", ".json"),
    recursivo: bool = False,
) -> list[Path]:
    """
    Remove evidencias cuja modificacao seja anterior ao limite informado.

    Args:
        diretorio (str | Path):
            Diretorio onde as evidencias serao pesquisadas.
        dias (int):
            Idade maxima permitida para os arquivos.
        extensoes (tuple[str, ...]):
            Extensoes consideradas na limpeza.
        recursivo (bool):
            Define se subdiretorios tambem serao pesquisados.

    Returns:
        list[Path]:
            Caminhos dos arquivos removidos com sucesso.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        ValueError:
            Caso dias seja negativo.
        OSError:
            Caso o sistema operacional nao consiga remover um arquivo.
    """
    # Valida a quantidade de dias antes de calcular o limite
    _validar_dias(dias)

    # Lista os arquivos que podem ser removidos
    capturas = listar_capturas(
        diretorio=diretorio,
        extensoes=extensoes,
        recursivo=recursivo,
    )

    # Calcula a data limite conforme a quantidade de dias
    data_limite = datetime.now() - timedelta(days=dias)
    timestamp_limite = data_limite.timestamp()

    # Inicializa a lista que registrara os arquivos removidos
    arquivos_removidos = []

    # Percorre as evidencias encontradas
    for arquivo in capturas:
        # Ignora arquivos mais recentes que a data limite
        if arquivo.stat().st_mtime >= timestamp_limite:
            continue

        # Remove o arquivo antigo
        arquivo.unlink()

        # Registra o caminho removido
        arquivos_removidos.append(arquivo)

    # Retorna a relacao dos arquivos removidos
    return arquivos_removidos
