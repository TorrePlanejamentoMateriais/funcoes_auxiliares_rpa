"""
============================================================
Módulo: SELENIUM / driver.py

Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para criação, configuração,
    validação e encerramento de drivers Selenium.
Navegadores suportados:
    - Google Chrome
    - Microsoft Edge
Funções disponíveis:
    - configurar_opcoes_chrome()
    - configurar_opcoes_edge()
    - criar_driver_chrome()
    - criar_driver_edge()
    - criar_driver()
    - verificar_driver_ativo()
    - obter_informacoes_driver()
    - fechar_aba_atual()
    - encerrar_driver()
    - reiniciar_driver()
Dependências:
    - selenium
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão de suporte ao Google Chrome e Microsoft Edge.
        - Inclusão de configurações de download, perfil e headless.
        - Inclusão de validações de caminhos e parâmetros.
        - Inclusão de funções para controle da sessão do driver.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa Path para manipulação e validação de caminhos
from pathlib import Path

# Importa Literal para restringir valores aceitos nas tipagens
from typing import Literal

# Importa os recursos principais do Selenium
from selenium import webdriver

# Importa as exceções tratadas pelo módulo
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    SessionNotCreatedException,
    WebDriverException,
)

# Importa as opções e os serviços dos navegadores suportados
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.remote.webdriver import WebDriver


# ----------------------------------------------------------------------------
# TIPOS
# ----------------------------------------------------------------------------

# Define os navegadores aceitos pelo módulo
Navegador = Literal["chrome", "edge"]

# Define as estratégias de carregamento aceitas pelo Selenium
EstrategiaCarregamento = Literal["normal", "eager", "none"]


# ----------------------------------------------------------------------------
# EXCEÇÕES PERSONALIZADAS
# ----------------------------------------------------------------------------

class DriverError(Exception):
    """Erro base para operações relacionadas ao WebDriver."""


class NavegadorNaoSuportadoError(DriverError):
    """Navegador informado não é suportado pela biblioteca."""


class DriverInicializacaoError(DriverError):
    """Não foi possível inicializar o WebDriver."""


class CaminhoDriverError(DriverError):
    """O caminho informado para driver, navegador ou perfil é inválido."""


# ----------------------------------------------------------------------------
# FUNÇÕES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------

def _normalizar_navegador(navegador: str) -> Navegador:
    """
    Normaliza e valida o nome do navegador.

    Args:
        navegador (str):
            Nome do navegador que será utilizado.

    Returns:
        Navegador:
            Nome padronizado como "chrome" ou "edge".

    Raises:
        TypeError:
            Caso o navegador informado não seja uma string.
        NavegadorNaoSuportadoError:
            Caso o navegador informado não seja suportado.
    """
    # Verifica se o navegador informado é uma string
    if not isinstance(navegador, str):
        raise TypeError(
            "O parâmetro 'navegador' deve ser uma string."
        )

    # Remove espaços e converte o nome para letras minúsculas
    navegador_normalizado = navegador.strip().lower()

    # Define os nomes aceitos e seus respectivos valores padronizados
    navegadores_aceitos = {
        "chrome": "chrome",
        "google": "chrome",
        "google chrome": "chrome",
        "edge": "edge",
        "microsoft edge": "edge",
        "msedge": "edge",
    }

    # Interrompe a execução quando o navegador não é suportado
    if navegador_normalizado not in navegadores_aceitos:
        raise NavegadorNaoSuportadoError(
            f"Navegador não suportado: '{navegador}'. "
            "Utilize 'chrome' ou 'edge'."
        )

    # Retorna o nome padronizado do navegador
    return navegadores_aceitos[navegador_normalizado]


def _validar_timeout(
    valor: int | float,
    nome: str
) -> None:
    """
    Valida um valor utilizado como timeout.

    Args:
        valor (int | float):
            Valor do timeout que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.

    Raises:
        TypeError:
            Caso o valor não seja numérico.
        ValueError:
            Caso o valor seja menor ou igual a zero.
    """
    # Verifica se o valor é numérico e rejeita booleanos
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise TypeError(
            f"O parâmetro '{nome}' deve ser numérico."
        )

    # Verifica se o valor é maior que zero
    if valor <= 0:
        raise ValueError(
            f"O parâmetro '{nome}' deve ser maior que zero."
        )


def _validar_estrategia_carregamento(
    estrategia: str
) -> EstrategiaCarregamento:
    """
    Valida a estratégia de carregamento das páginas.

    Args:
        estrategia (str):
            Estratégia normal, eager ou none.

    Returns:
        EstrategiaCarregamento:
            Estratégia normalizada.

    Raises:
        TypeError:
            Caso a estratégia não seja uma string.
        ValueError:
            Caso a estratégia não seja reconhecida.
    """
    # Verifica se a estratégia informada é uma string
    if not isinstance(estrategia, str):
        raise TypeError(
            "O parâmetro 'estrategia_carregamento' deve ser uma string."
        )

    # Remove espaços e converte o valor para letras minúsculas
    estrategia_normalizada = estrategia.strip().lower()

    # Define as estratégias permitidas
    estrategias_validas = {"normal", "eager", "none"}

    # Interrompe a execução quando a estratégia é inválida
    if estrategia_normalizada not in estrategias_validas:
        raise ValueError(
            "Estratégia de carregamento inválida. "
            "Utilize 'normal', 'eager' ou 'none'."
        )

    # Retorna a estratégia normalizada
    return estrategia_normalizada


def _preparar_diretorio_download(
    diretorio_download: str | Path | None
) -> str | None:
    """
    Cria e normaliza o diretório utilizado para downloads.

    Args:
        diretorio_download (str | Path | None):
            Diretório onde os arquivos serão baixados.

    Returns:
        str | None:
            Caminho absoluto do diretório ou None.

    Raises:
        TypeError:
            Caso o caminho não seja uma string, Path ou None.
        OSError:
            Caso não seja possível criar o diretório.
    """
    # Retorna None quando nenhum diretório é informado
    if diretorio_download is None:
        return None

    # Verifica o tipo do caminho informado
    if not isinstance(diretorio_download, (str, Path)):
        raise TypeError(
            "O parâmetro 'diretorio_download' deve ser "
            "uma string, Path ou None."
        )

    # Converte o caminho para um caminho absoluto
    diretorio = Path(diretorio_download).expanduser().resolve()

    # Cria o diretório e seus diretórios pais quando necessário
    diretorio.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Retorna o caminho absoluto como string
    return str(diretorio)


def _validar_caminho_arquivo(
    caminho: str | Path | None,
    nome_parametro: str
) -> str | None:
    """
    Valida um caminho opcional para um arquivo.

    Args:
        caminho (str | Path | None):
            Caminho do arquivo que será validado.
        nome_parametro (str):
            Nome do parâmetro utilizado na mensagem de erro.

    Returns:
        str | None:
            Caminho absoluto do arquivo ou None.

    Raises:
        TypeError:
            Caso o caminho não seja uma string, Path ou None.
        CaminhoDriverError:
            Caso o arquivo informado não exista.
    """
    # Retorna None quando nenhum arquivo é informado
    if caminho is None:
        return None

    # Verifica o tipo do caminho informado
    if not isinstance(caminho, (str, Path)):
        raise TypeError(
            f"O parâmetro '{nome_parametro}' deve ser "
            "uma string, Path ou None."
        )

    # Converte o caminho para um caminho absoluto
    caminho_validado = Path(caminho).expanduser().resolve()

    # Interrompe a execução quando o arquivo não existe
    if not caminho_validado.is_file():
        raise CaminhoDriverError(
            f"O arquivo informado em '{nome_parametro}' "
            f"não foi encontrado: {caminho_validado}"
        )

    # Retorna o caminho absoluto como string
    return str(caminho_validado)


def _validar_diretorio_existente(
    caminho: str | Path | None,
    nome_parametro: str
) -> str | None:
    """
    Valida um caminho opcional para um diretório existente.

    Args:
        caminho (str | Path | None):
            Caminho do diretório que será validado.
        nome_parametro (str):
            Nome do parâmetro utilizado na mensagem de erro.

    Returns:
        str | None:
            Caminho absoluto do diretório ou None.

    Raises:
        TypeError:
            Caso o caminho não seja uma string, Path ou None.
        CaminhoDriverError:
            Caso o diretório informado não exista.
    """
    # Retorna None quando nenhum diretório é informado
    if caminho is None:
        return None

    # Verifica o tipo do caminho informado
    if not isinstance(caminho, (str, Path)):
        raise TypeError(
            f"O parâmetro '{nome_parametro}' deve ser "
            "uma string, Path ou None."
        )

    # Converte o caminho para um caminho absoluto
    diretorio = Path(caminho).expanduser().resolve()

    # Interrompe a execução quando o diretório não existe
    if not diretorio.is_dir():
        raise CaminhoDriverError(
            f"O diretório informado em '{nome_parametro}' "
            f"não foi encontrado: {diretorio}"
        )

    # Retorna o caminho absoluto como string
    return str(diretorio)


def _obter_preferencias_download(
    diretorio_download: str | None,
    permitir_multiplos_downloads: bool
) -> dict:
    """
    Retorna as preferências utilizadas para downloads.

    Args:
        diretorio_download (str | None):
            Caminho absoluto do diretório de download.
        permitir_multiplos_downloads (bool):
            Permite vários downloads automáticos.

    Returns:
        dict:
            Preferências que serão aplicadas ao navegador.
    """
    # Define as preferências padrão do navegador
    preferencias = {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }

    # Adiciona o diretório padrão quando ele é informado
    if diretorio_download is not None:
        preferencias["download.default_directory"] = diretorio_download

    # Permite múltiplos downloads automáticos quando solicitado
    if permitir_multiplos_downloads:
        preferencias[
            "profile.default_content_setting_values.automatic_downloads"
        ] = 1

    # Retorna as preferências configuradas
    return preferencias


def _adicionar_argumentos_comuns(
    options: ChromeOptions | EdgeOptions,
    navegador: Navegador,
    headless: bool,
    maximizado: bool,
    tamanho_janela: tuple[int, int] | None,
    modo_privado: bool,
    ignorar_certificados: bool,
    desabilitar_notificacoes: bool,
    desabilitar_gpu: bool,
    desabilitar_sandbox: bool,
    user_agent: str | None,
    proxy: str | None,
    diretorio_perfil: str | None,
    nome_perfil: str | None,
    argumentos_adicionais: list[str] | None
) -> None:
    """
    Adiciona argumentos compartilhados pelo Chrome e pelo Edge.

    Args:
        options (ChromeOptions | EdgeOptions):
            Objeto de opções que receberá os argumentos.
        navegador (Navegador):
            Navegador que está sendo configurado.
        headless (bool):
            Executa o navegador sem interface gráfica.
        maximizado (bool):
            Inicia o navegador maximizado.
        tamanho_janela (tuple[int, int] | None):
            Largura e altura utilizadas na janela.
        modo_privado (bool):
            Ativa o modo anônimo ou InPrivate.
        ignorar_certificados (bool):
            Ignora erros de certificados inseguros.
        desabilitar_notificacoes (bool):
            Desabilita notificações do navegador.
        desabilitar_gpu (bool):
            Desabilita a aceleração por GPU.
        desabilitar_sandbox (bool):
            Desabilita o sandbox e o uso de memória compartilhada.
        user_agent (str | None):
            User-Agent personalizado.
        proxy (str | None):
            Endereço do servidor proxy.
        diretorio_perfil (str | None):
            Diretório raiz dos dados do usuário.
        nome_perfil (str | None):
            Nome do perfil que será utilizado.
        argumentos_adicionais (list[str] | None):
            Argumentos extras enviados ao navegador.

    Raises:
        TypeError:
            Caso algum argumento possua tipo inválido.
        ValueError:
            Caso dimensões ou strings obrigatórias sejam inválidas.
    """
    if headless:
        options.add_argument("--headless=new")

    if maximizado and not headless:
        options.add_argument("--start-maximized")

    if tamanho_janela is not None:
        if not isinstance(tamanho_janela, tuple) or len(tamanho_janela) != 2:
            raise TypeError(
                "O parâmetro 'tamanho_janela' deve ser uma tupla "
                "contendo largura e altura."
            )

        largura, altura = tamanho_janela

        if (
            isinstance(largura, bool)
            or isinstance(altura, bool)
            or not isinstance(largura, int)
            or not isinstance(altura, int)
        ):
            raise TypeError(
                "A largura e a altura da janela devem ser números inteiros."
            )

        if largura <= 0 or altura <= 0:
            raise ValueError(
                "A largura e a altura da janela devem ser maiores que zero."
            )

        options.add_argument(f"--window-size={largura},{altura}")

    if modo_privado:
        argumento_privado = "--incognito" if navegador == "chrome" else "--inprivate"
        options.add_argument(argumento_privado)

    if ignorar_certificados:
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--allow-insecure-localhost")

    if desabilitar_notificacoes:
        options.add_argument("--disable-notifications")

    if desabilitar_gpu:
        options.add_argument("--disable-gpu")

    if desabilitar_sandbox:
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

    if user_agent is not None:
        if not isinstance(user_agent, str):
            raise TypeError(
                "O parâmetro 'user_agent' deve ser uma string ou None."
            )
        if not user_agent.strip():
            raise ValueError(
                "O parâmetro 'user_agent' não pode ser uma string vazia."
            )
        options.add_argument(f"--user-agent={user_agent.strip()}")

    if proxy is not None:
        if not isinstance(proxy, str):
            raise TypeError(
                "O parâmetro 'proxy' deve ser uma string ou None."
            )
        if not proxy.strip():
            raise ValueError(
                "O parâmetro 'proxy' não pode ser uma string vazia."
            )
        options.add_argument(f"--proxy-server={proxy.strip()}")

    if diretorio_perfil is not None:
        options.add_argument(f"--user-data-dir={diretorio_perfil}")

    if nome_perfil is not None:
        if not isinstance(nome_perfil, str):
            raise TypeError(
                "O parâmetro 'nome_perfil' deve ser uma string ou None."
            )
        if not nome_perfil.strip():
            raise ValueError(
                "O parâmetro 'nome_perfil' não pode ser vazio."
            )
        options.add_argument(f"--profile-directory={nome_perfil.strip()}")

    if argumentos_adicionais is not None:
        if not isinstance(argumentos_adicionais, list):
            raise TypeError(
                "O parâmetro 'argumentos_adicionais' deve ser "
                "uma lista de strings ou None."
            )

        if not all(isinstance(argumento, str) for argumento in argumentos_adicionais):
            raise TypeError(
                "Todos os argumentos adicionais devem ser strings."
            )

        for argumento in argumentos_adicionais:
            if argumento.strip():
                options.add_argument(argumento.strip())


# ----------------------------------------------------------------------------
# FUNÇÕES DE CONFIGURAÇÃO
# ----------------------------------------------------------------------------

def configurar_opcoes_chrome(
    headless: bool = False,
    diretorio_download: str | Path | None = None,
    maximizado: bool = True,
    tamanho_janela: tuple[int, int] | None = None,
    modo_privado: bool = False,
    ignorar_certificados: bool = True,
    desabilitar_notificacoes: bool = True,
    desabilitar_gpu: bool = False,
    desabilitar_sandbox: bool = False,
    manter_aberto: bool = False,
    permitir_multiplos_downloads: bool = True,
    estrategia_carregamento: EstrategiaCarregamento = "normal",
    user_agent: str | None = None,
    proxy: str | None = None,
    diretorio_perfil: str | Path | None = None,
    nome_perfil: str | None = None,
    caminho_navegador: str | Path | None = None,
    extensoes: list[str | Path] | None = None,
    argumentos_adicionais: list[str] | None = None
) -> ChromeOptions:
    """
    Configura e retorna as opções do Google Chrome.

    Args:
        headless (bool):
            Executa o Chrome sem abrir uma janela visível.
        diretorio_download (str | Path | None):
            Diretório utilizado para salvar os downloads.
        maximizado (bool):
            Inicia o navegador maximizado.
        tamanho_janela (tuple[int, int] | None):
            Define a largura e a altura da janela.
        modo_privado (bool):
            Inicia o navegador no modo anônimo.
        ignorar_certificados (bool):
            Ignora erros de certificados inseguros.
        desabilitar_notificacoes (bool):
            Desabilita notificações do navegador.
        desabilitar_gpu (bool):
            Desabilita a aceleração por GPU.
        desabilitar_sandbox (bool):
            Adiciona argumentos úteis para ambientes sem sandbox.
        manter_aberto (bool):
            Mantém o navegador aberto ao final do processo quando
            driver.quit() não for executado.
        permitir_multiplos_downloads (bool):
            Permite múltiplos downloads automáticos.
        estrategia_carregamento (EstrategiaCarregamento):
            Estratégia normal, eager ou none.
        user_agent (str | None):
            User-Agent personalizado.
        proxy (str | None):
            Servidor proxy utilizado pelo navegador.
        diretorio_perfil (str | Path | None):
            Diretório raiz dos dados do usuário do Chrome.
        nome_perfil (str | None):
            Nome do perfil, como Default ou Profile 1.
        caminho_navegador (str | Path | None):
            Caminho opcional do executável chrome.exe.
        extensoes (list[str | Path] | None):
            Lista de extensões no formato CRX.
        argumentos_adicionais (list[str] | None):
            Argumentos adicionais enviados ao Chrome.

    Returns:
        ChromeOptions:
            Objeto contendo as opções configuradas do Chrome.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
        CaminhoDriverError:
            Caso um caminho informado não exista.
    """
    if not isinstance(headless, bool):
        raise TypeError("O parâmetro 'headless' deve ser booleano.")

    diretorio_download_validado = _preparar_diretorio_download(
        diretorio_download
    )
    diretorio_perfil_validado = _validar_diretorio_existente(
        diretorio_perfil,
        "diretorio_perfil",
    )
    caminho_navegador_validado = _validar_caminho_arquivo(
        caminho_navegador,
        "caminho_navegador",
    )
    estrategia_validada = _validar_estrategia_carregamento(
        estrategia_carregamento
    )

    options = ChromeOptions()
    options.page_load_strategy = estrategia_validada
    options.add_experimental_option(
        "prefs",
        _obter_preferencias_download(
            diretorio_download_validado,
            permitir_multiplos_downloads,
        ),
    )
    options.add_experimental_option(
        "excludeSwitches",
        ["enable-automation", "enable-logging"],
    )
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("detach", manter_aberto)

    if caminho_navegador_validado is not None:
        options.binary_location = caminho_navegador_validado

    _adicionar_argumentos_comuns(
        options=options,
        navegador="chrome",
        headless=headless,
        maximizado=maximizado,
        tamanho_janela=tamanho_janela,
        modo_privado=modo_privado,
        ignorar_certificados=ignorar_certificados,
        desabilitar_notificacoes=desabilitar_notificacoes,
        desabilitar_gpu=desabilitar_gpu,
        desabilitar_sandbox=desabilitar_sandbox,
        user_agent=user_agent,
        proxy=proxy,
        diretorio_perfil=diretorio_perfil_validado,
        nome_perfil=nome_perfil,
        argumentos_adicionais=argumentos_adicionais,
    )

    if extensoes is not None:
        if not isinstance(extensoes, list):
            raise TypeError(
                "O parâmetro 'extensoes' deve ser uma lista ou None."
            )
        for extensao in extensoes:
            caminho_extensao = _validar_caminho_arquivo(
                extensao,
                "extensoes",
            )
            options.add_extension(caminho_extensao)

    return options


def configurar_opcoes_edge(
    headless: bool = False,
    diretorio_download: str | Path | None = None,
    maximizado: bool = True,
    tamanho_janela: tuple[int, int] | None = None,
    modo_privado: bool = False,
    ignorar_certificados: bool = True,
    desabilitar_notificacoes: bool = True,
    desabilitar_gpu: bool = False,
    desabilitar_sandbox: bool = False,
    manter_aberto: bool = False,
    permitir_multiplos_downloads: bool = True,
    estrategia_carregamento: EstrategiaCarregamento = "normal",
    user_agent: str | None = None,
    proxy: str | None = None,
    diretorio_perfil: str | Path | None = None,
    nome_perfil: str | None = None,
    caminho_navegador: str | Path | None = None,
    extensoes: list[str | Path] | None = None,
    argumentos_adicionais: list[str] | None = None
) -> EdgeOptions:
    """
    Configura e retorna as opções do Microsoft Edge.

    Args:
        headless (bool):
            Executa o Edge sem abrir uma janela visível.
        diretorio_download (str | Path | None):
            Diretório utilizado para salvar os downloads.
        maximizado (bool):
            Inicia o navegador maximizado.
        tamanho_janela (tuple[int, int] | None):
            Define a largura e a altura da janela.
        modo_privado (bool):
            Inicia o navegador no modo InPrivate.
        ignorar_certificados (bool):
            Ignora erros de certificados inseguros.
        desabilitar_notificacoes (bool):
            Desabilita notificações do navegador.
        desabilitar_gpu (bool):
            Desabilita a aceleração por GPU.
        desabilitar_sandbox (bool):
            Adiciona argumentos úteis para ambientes sem sandbox.
        manter_aberto (bool):
            Mantém o navegador aberto ao final do processo quando
            driver.quit() não for executado.
        permitir_multiplos_downloads (bool):
            Permite múltiplos downloads automáticos.
        estrategia_carregamento (EstrategiaCarregamento):
            Estratégia normal, eager ou none.
        user_agent (str | None):
            User-Agent personalizado.
        proxy (str | None):
            Servidor proxy utilizado pelo navegador.
        diretorio_perfil (str | Path | None):
            Diretório raiz dos dados do usuário do Edge.
        nome_perfil (str | None):
            Nome do perfil, como Default ou Profile 1.
        caminho_navegador (str | Path | None):
            Caminho opcional do executável msedge.exe.
        extensoes (list[str | Path] | None):
            Lista de extensões no formato CRX.
        argumentos_adicionais (list[str] | None):
            Argumentos adicionais enviados ao Edge.

    Returns:
        EdgeOptions:
            Objeto contendo as opções configuradas do Edge.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo inválido.
        ValueError:
            Caso algum valor informado seja inválido.
        CaminhoDriverError:
            Caso um caminho informado não exista.
    """
    if not isinstance(headless, bool):
        raise TypeError("O parâmetro 'headless' deve ser booleano.")

    diretorio_download_validado = _preparar_diretorio_download(
        diretorio_download
    )
    diretorio_perfil_validado = _validar_diretorio_existente(
        diretorio_perfil,
        "diretorio_perfil",
    )
    caminho_navegador_validado = _validar_caminho_arquivo(
        caminho_navegador,
        "caminho_navegador",
    )
    estrategia_validada = _validar_estrategia_carregamento(
        estrategia_carregamento
    )

    options = EdgeOptions()
    options.page_load_strategy = estrategia_validada
    options.add_experimental_option(
        "prefs",
        _obter_preferencias_download(
            diretorio_download_validado,
            permitir_multiplos_downloads,
        ),
    )
    options.add_experimental_option(
        "excludeSwitches",
        ["enable-automation", "enable-logging"],
    )
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("detach", manter_aberto)

    if caminho_navegador_validado is not None:
        options.binary_location = caminho_navegador_validado

    _adicionar_argumentos_comuns(
        options=options,
        navegador="edge",
        headless=headless,
        maximizado=maximizado,
        tamanho_janela=tamanho_janela,
        modo_privado=modo_privado,
        ignorar_certificados=ignorar_certificados,
        desabilitar_notificacoes=desabilitar_notificacoes,
        desabilitar_gpu=desabilitar_gpu,
        desabilitar_sandbox=desabilitar_sandbox,
        user_agent=user_agent,
        proxy=proxy,
        diretorio_perfil=diretorio_perfil_validado,
        nome_perfil=nome_perfil,
        argumentos_adicionais=argumentos_adicionais,
    )

    if extensoes is not None:
        if not isinstance(extensoes, list):
            raise TypeError(
                "O parâmetro 'extensoes' deve ser uma lista ou None."
            )
        for extensao in extensoes:
            caminho_extensao = _validar_caminho_arquivo(
                extensao,
                "extensoes",
            )
            options.add_extension(caminho_extensao)

    return options


# ----------------------------------------------------------------------------
# FUNÇÕES DE INICIALIZAÇÃO
# ----------------------------------------------------------------------------

def criar_driver_chrome(
    caminho_driver: str | Path | None = None,
    timeout_pagina: int | float = 60,
    timeout_script: int | float = 30,
    espera_implicita: int | float = 0,
    **configuracoes
) -> WebDriver:
    """
    Cria e retorna uma instância configurada do Google Chrome.

    Quando o caminho do driver não é informado, o Selenium Manager
    gerencia automaticamente o ChromeDriver.

    Args:
        caminho_driver (str | Path | None):
            Caminho opcional para o executável chromedriver.exe.
        timeout_pagina (int | float):
            Tempo máximo para carregamento da página.
        timeout_script (int | float):
            Tempo máximo para execução de scripts assíncronos.
        espera_implicita (int | float):
            Tempo utilizado como espera implícita.
        **configuracoes:
            Configurações encaminhadas para configurar_opcoes_chrome().

    Returns:
        WebDriver:
            Instância configurada do Google Chrome.

    Raises:
        TypeError:
            Caso algum timeout ou espera possua tipo inválido.
        ValueError:
            Caso algum timeout ou espera possua valor inválido.
        CaminhoDriverError:
            Caso o caminho manual do driver não exista.
        DriverInicializacaoError:
            Caso o navegador não possa ser iniciado.
    """
    _validar_timeout(timeout_pagina, "timeout_pagina")
    _validar_timeout(timeout_script, "timeout_script")

    if (
        isinstance(espera_implicita, bool)
        or not isinstance(espera_implicita, (int, float))
    ):
        raise TypeError(
            "O parâmetro 'espera_implicita' deve ser numérico."
        )

    if espera_implicita < 0:
        raise ValueError(
            "O parâmetro 'espera_implicita' não pode ser negativo."
        )

    caminho_driver_validado = _validar_caminho_arquivo(
        caminho_driver,
        "caminho_driver",
    )

    try:
        options = configurar_opcoes_chrome(**configuracoes)

        if caminho_driver_validado is not None:
            service = ChromeService(
                executable_path=caminho_driver_validado
            )
            driver = webdriver.Chrome(
                service=service,
                options=options,
            )
        else:
            driver = webdriver.Chrome(options=options)

        driver.set_page_load_timeout(timeout_pagina)
        driver.set_script_timeout(timeout_script)
        driver.implicitly_wait(espera_implicita)

        return driver

    except SessionNotCreatedException as error:
        raise DriverInicializacaoError(
            "Não foi possível criar a sessão do Google Chrome. "
            "Verifique o navegador e a compatibilidade do driver."
        ) from error

    except WebDriverException as error:
        raise DriverInicializacaoError(
            "Erro ao inicializar o Google Chrome. "
            f"Detalhes: {error}"
        ) from error


def criar_driver_edge(
    caminho_driver: str | Path | None = None,
    timeout_pagina: int | float = 60,
    timeout_script: int | float = 30,
    espera_implicita: int | float = 0,
    **configuracoes
) -> WebDriver:
    """
    Cria e retorna uma instância configurada do Microsoft Edge.

    Quando o caminho do driver não é informado, o Selenium Manager
    gerencia automaticamente o EdgeDriver.

    Args:
        caminho_driver (str | Path | None):
            Caminho opcional para o executável msedgedriver.exe.
        timeout_pagina (int | float):
            Tempo máximo para carregamento da página.
        timeout_script (int | float):
            Tempo máximo para execução de scripts assíncronos.
        espera_implicita (int | float):
            Tempo utilizado como espera implícita.
        **configuracoes:
            Configurações encaminhadas para configurar_opcoes_edge().

    Returns:
        WebDriver:
            Instância configurada do Microsoft Edge.

    Raises:
        TypeError:
            Caso algum timeout ou espera possua tipo inválido.
        ValueError:
            Caso algum timeout ou espera possua valor inválido.
        CaminhoDriverError:
            Caso o caminho manual do driver não exista.
        DriverInicializacaoError:
            Caso o navegador não possa ser iniciado.
    """
    _validar_timeout(timeout_pagina, "timeout_pagina")
    _validar_timeout(timeout_script, "timeout_script")

    if (
        isinstance(espera_implicita, bool)
        or not isinstance(espera_implicita, (int, float))
    ):
        raise TypeError(
            "O parâmetro 'espera_implicita' deve ser numérico."
        )

    if espera_implicita < 0:
        raise ValueError(
            "O parâmetro 'espera_implicita' não pode ser negativo."
        )

    caminho_driver_validado = _validar_caminho_arquivo(
        caminho_driver,
        "caminho_driver",
    )

    try:
        options = configurar_opcoes_edge(**configuracoes)

        if caminho_driver_validado is not None:
            service = EdgeService(
                executable_path=caminho_driver_validado
            )
            driver = webdriver.Edge(
                service=service,
                options=options,
            )
        else:
            driver = webdriver.Edge(options=options)

        driver.set_page_load_timeout(timeout_pagina)
        driver.set_script_timeout(timeout_script)
        driver.implicitly_wait(espera_implicita)

        return driver

    except SessionNotCreatedException as error:
        raise DriverInicializacaoError(
            "Não foi possível criar a sessão do Microsoft Edge. "
            "Verifique o navegador e a compatibilidade do driver."
        ) from error

    except WebDriverException as error:
        raise DriverInicializacaoError(
            "Erro ao inicializar o Microsoft Edge. "
            f"Detalhes: {error}"
        ) from error


def criar_driver(
    navegador: str = "edge",
    caminho_driver: str | Path | None = None,
    timeout_pagina: int | float = 60,
    timeout_script: int | float = 30,
    espera_implicita: int | float = 0,
    **configuracoes
) -> WebDriver:
    """
    Cria um WebDriver para Google Chrome ou Microsoft Edge.

    Esta é a função principal recomendada para os projetos de RPA.

    Args:
        navegador (str):
            Navegador utilizado pela automação.
        caminho_driver (str | Path | None):
            Caminho opcional para o executável do driver.
        timeout_pagina (int | float):
            Tempo máximo para carregamento das páginas.
        timeout_script (int | float):
            Tempo máximo para scripts assíncronos.
        espera_implicita (int | float):
            Tempo utilizado como espera implícita.
        **configuracoes:
            Configurações específicas do navegador.

    Returns:
        WebDriver:
            Instância configurada do navegador selecionado.

    Raises:
        NavegadorNaoSuportadoError:
            Caso o navegador não seja Chrome ou Edge.
        DriverInicializacaoError:
            Caso não seja possível inicializar o navegador.
    """
    navegador_normalizado = _normalizar_navegador(navegador)

    if navegador_normalizado == "chrome":
        return criar_driver_chrome(
            caminho_driver=caminho_driver,
            timeout_pagina=timeout_pagina,
            timeout_script=timeout_script,
            espera_implicita=espera_implicita,
            **configuracoes,
        )

    return criar_driver_edge(
        caminho_driver=caminho_driver,
        timeout_pagina=timeout_pagina,
        timeout_script=timeout_script,
        espera_implicita=espera_implicita,
        **configuracoes,
    )


# ----------------------------------------------------------------------------
# FUNÇÕES DE CONTROLE DA SESSÃO
# ----------------------------------------------------------------------------

def verificar_driver_ativo(
    driver: WebDriver | None
) -> bool:
    """
    Verifica se a sessão do WebDriver ainda está ativa.

    Args:
        driver (WebDriver | None):
            Instância que será verificada.

    Returns:
        bool:
            True quando a sessão estiver ativa e False caso contrário.
    """
    if driver is None:
        return False

    try:
        driver.current_url
        return True
    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ):
        return False


def obter_informacoes_driver(
    driver: WebDriver
) -> dict:
    """
    Retorna informações sobre o navegador e a sessão atual.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.

    Returns:
        dict:
            Navegador, versão, plataforma, sessão, URL, título e
            quantidade de janelas abertas.

    Raises:
        DriverError:
            Caso o WebDriver não esteja ativo.
    """
    if not verificar_driver_ativo(driver):
        raise DriverError(
            "Não foi possível obter as informações porque "
            "o WebDriver não está ativo."
        )

    capabilities = driver.capabilities

    return {
        "navegador": capabilities.get("browserName"),
        "versao_navegador": capabilities.get("browserVersion"),
        "plataforma": capabilities.get("platformName"),
        "session_id": driver.session_id,
        "url_atual": driver.current_url,
        "titulo_atual": driver.title,
        "quantidade_janelas": len(driver.window_handles),
    }


def fechar_aba_atual(
    driver: WebDriver,
    trocar_para_ultima_aba: bool = True
) -> None:
    """
    Fecha a aba atual e, opcionalmente, seleciona a última aba.

    Args:
        driver (WebDriver):
            Instância ativa do Selenium WebDriver.
        trocar_para_ultima_aba (bool):
            Define se o foco será transferido para a última aba aberta.

    Raises:
        DriverError:
            Caso o WebDriver não esteja ativo.
    """
    if not verificar_driver_ativo(driver):
        raise DriverError(
            "Não foi possível fechar a aba porque "
            "o WebDriver não está ativo."
        )

    driver.close()

    if not trocar_para_ultima_aba:
        return

    try:
        janelas_disponiveis = driver.window_handles

        if janelas_disponiveis:
            driver.switch_to.window(janelas_disponiveis[-1])
    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ):
        return


def encerrar_driver(
    driver: WebDriver | None,
    ignorar_erros: bool = True
) -> None:
    """
    Encerra todas as janelas e finaliza a sessão do WebDriver.

    Args:
        driver (WebDriver | None):
            Instância que será encerrada.
        ignorar_erros (bool):
            Define se erros de encerramento serão ignorados.

    Raises:
        DriverError:
            Caso ocorra uma falha e ignorar_erros seja False.
    """
    if driver is None:
        return

    try:
        driver.quit()
    except (
        InvalidSessionIdException,
        NoSuchWindowException,
        WebDriverException,
    ) as error:
        if not ignorar_erros:
            raise DriverError(
                "Ocorreu um erro ao encerrar o WebDriver."
            ) from error


def reiniciar_driver(
    driver: WebDriver | None = None,
    navegador: str = "edge",
    **configuracoes
) -> WebDriver:
    """
    Encerra o driver atual e cria uma nova sessão.

    Args:
        driver (WebDriver | None):
            Instância atual que será encerrada.
        navegador (str):
            Navegador utilizado na nova sessão.
        **configuracoes:
            Configurações encaminhadas para criar_driver().

    Returns:
        WebDriver:
            Nova instância configurada do navegador.
    """
    encerrar_driver(
        driver=driver,
        ignorar_erros=True,
    )

    return criar_driver(
        navegador=navegador,
        **configuracoes,
    )
