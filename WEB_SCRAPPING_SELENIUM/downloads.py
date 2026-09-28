"""
============================================================
Modulo: SELENIUM / downloads.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para configuracao,
    acompanhamento, validacao, organizacao e limpeza de downloads
    realizados por automacoes Selenium.
Funcoes disponiveis:
    - preparar_diretorio_downloads()
    - configurar_downloads_chromium()
    - listar_arquivos_download()
    - obter_snapshot_downloads()
    - identificar_novos_arquivos()
    - existe_download_em_andamento()
    - verificar_download_concluido()
    - aguardar_inicio_download()
    - aguardar_download()
    - aguardar_novos_downloads()
    - aguardar_downloads_concluidos()
    - aguardar_arquivo_estavel()
    - obter_arquivo_mais_recente()
    - localizar_arquivo_download()
    - validar_arquivo_download()
    - calcular_hash_arquivo()
    - renomear_download()
    - mover_download()
    - copiar_download()
    - remover_download()
    - limpar_diretorio_downloads()
    - clicar_e_aguardar_download()
Dependencias:
    - selenium
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao do monitoramento de arquivos temporarios.
        - Inclusao da espera por downloads unicos e multiplos.
        - Inclusao de validacao de tamanho, extensao e estabilidade.
        - Inclusao de movimentacao, copia, renomeacao e limpeza.
        - Inclusao de integracao entre clique Selenium e download.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa hashlib para calcular hashes de integridade dos arquivos
import hashlib

# Importa shutil para copiar e mover arquivos entre diretorios
import shutil

# Importa time para controlar as verificacoes periodicas dos downloads
import time

# Importa Path para manipular caminhos de arquivos e diretorios
from pathlib import Path

# Importa Literal para restringir os algoritmos de hash aceitos
from typing import Literal

# Importa excecoes utilizadas no tratamento das interacoes Selenium
from selenium.common.exceptions import (
    InvalidSessionIdException,
    NoSuchWindowException,
    TimeoutException,
    WebDriverException,
)

# Importa os tipos principais do Selenium
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

# Importa as condicoes e a espera explicita para o clique de download
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ----------------------------------------------------------------------------
# TIPOS E CONSTANTES
# ----------------------------------------------------------------------------

# Define o formato padrao de um localizador Selenium
Localizador = tuple[str, str]

# Define os algoritmos de hash aceitos pelo modulo
AlgoritmoHash = Literal["md5", "sha1", "sha256", "sha512"]

# Define as extensoes utilizadas por navegadores durante downloads incompletos
EXTENSOES_TEMPORARIAS_PADRAO = (
    ".crdownload",
    ".part",
    ".partial",
    ".tmp",
    ".download",
)


# ----------------------------------------------------------------------------
# EXCECOES PERSONALIZADAS
# ----------------------------------------------------------------------------

class DownloadsError(Exception):
    """Erro base para operacoes relacionadas a downloads."""


class DriverInativoError(DownloadsError):
    """O WebDriver informado nao possui uma sessao ativa."""


class LocalizadorInvalidoError(DownloadsError):
    """O localizador Selenium informado nao possui formato valido."""


class DiretorioDownloadError(DownloadsError):
    """O diretorio informado para downloads nao e valido."""


class ArquivoDownloadNaoEncontradoError(DownloadsError):
    """O arquivo de download solicitado nao foi encontrado."""


class DownloadTimeoutError(DownloadsError):
    """O download nao iniciou ou nao terminou dentro do tempo limite."""


class DownloadIncompletoError(DownloadsError):
    """O arquivo ainda esta incompleto ou nao atende aos criterios definidos."""


class ArquivoDownloadExistenteError(DownloadsError):
    """Ja existe um arquivo no caminho de destino informado."""


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
    # Verifica se o objeto recebido e uma instancia de WebDriver
    if not isinstance(driver, WebDriver):
        raise TypeError(
            "O parametro 'driver' deve ser uma instancia de WebDriver."
        )

    # Interrompe a execucao quando nao existe identificador de sessao
    if not driver.session_id:
        raise DriverInativoError(
            "O WebDriver informado nao possui uma sessao ativa."
        )

    # Confirma que a sessao responde ao acesso da janela atual
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
            Caso o localizador nao possua duas strings preenchidas.
    """
    # Verifica se o localizador e uma tupla com exatamente dois itens
    if not isinstance(localizador, tuple) or len(localizador) != 2:
        raise LocalizadorInvalidoError(
            "O parametro 'localizador' deve ser uma tupla com dois itens, "
            "por exemplo: (By.ID, 'botao_download')."
        )

    # Separa os itens para validar seu conteudo
    estrategia, valor = localizador

    # Verifica se a estrategia e o valor sao strings nao vazias
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


def _validar_numero_positivo(
    valor: int | float,
    nome_parametro: str,
    permitir_zero: bool = False,
) -> None:
    """
    Valida um numero positivo utilizado pelo modulo.

    Args:
        valor (int | float):
            Valor numerico que sera validado.
        nome_parametro (str):
            Nome exibido nas mensagens de erro.
        permitir_zero (bool):
            Define se o valor zero sera considerado valido.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso o valor nao seja numerico ou permitir_zero nao seja booleano.
        ValueError:
            Caso o valor esteja fora do intervalo permitido.
    """
    # Valida a opcao booleana antes de avaliar o numero
    if not isinstance(permitir_zero, bool):
        raise TypeError(
            "O parametro 'permitir_zero' deve ser booleano."
        )

    # Rejeita booleanos e objetos que nao sejam numeros
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise TypeError(
            f"O parametro '{nome_parametro}' deve ser numerico."
        )

    # Define a regra de limite conforme a permissao de zero
    valor_invalido = valor < 0 if permitir_zero else valor <= 0

    # Interrompe a execucao quando o valor nao atende ao limite
    if valor_invalido:
        operador = "maior ou igual a zero" if permitir_zero else "maior que zero"
        raise ValueError(
            f"O parametro '{nome_parametro}' deve ser {operador}."
        )


def _normalizar_extensoes(
    extensoes: tuple[str, ...] | list[str] | str | None,
) -> tuple[str, ...] | None:
    """
    Normaliza uma extensao ou colecao de extensoes.

    Args:
        extensoes (tuple[str, ...] | list[str] | str | None):
            Extensoes que serao utilizadas em filtros de arquivos.

    Returns:
        tuple[str, ...] | None:
            Extensoes em letras minusculas e iniciadas por ponto.

    Raises:
        TypeError:
            Caso o parametro ou algum item possua tipo invalido.
        ValueError:
            Caso a colecao esteja vazia ou contenha string vazia.
    """
    # Mantem None quando nenhum filtro de extensao foi solicitado
    if extensoes is None:
        return None

    # Converte uma extensao unica para uma tupla
    if isinstance(extensoes, str):
        extensoes = (extensoes,)

    # Verifica se o valor final e uma lista ou tupla
    if not isinstance(extensoes, (list, tuple)):
        raise TypeError(
            "O parametro 'extensoes' deve ser string, lista, tupla ou None."
        )

    # Interrompe a execucao quando a colecao esta vazia
    if not extensoes:
        raise ValueError(
            "O parametro 'extensoes' nao pode estar vazio."
        )

    # Inicializa a lista de extensoes tratadas
    resultado = []

    # Percorre e normaliza cada extensao informada
    for extensao in extensoes:
        if not isinstance(extensao, str):
            raise TypeError(
                "Todos os itens de 'extensoes' devem ser strings."
            )

        extensao_tratada = extensao.strip().lower()

        if not extensao_tratada:
            raise ValueError(
                "As extensoes nao podem possuir valores vazios."
            )

        if not extensao_tratada.startswith("."):
            extensao_tratada = f".{extensao_tratada}"

        resultado.append(extensao_tratada)

    # Remove duplicidades preservando a ordem
    return tuple(dict.fromkeys(resultado))


def _validar_caminho_arquivo(
    caminho: str | Path,
    deve_existir: bool = True,
) -> Path:
    """
    Normaliza e valida um caminho de arquivo.

    Args:
        caminho (str | Path):
            Caminho do arquivo que sera validado.
        deve_existir (bool):
            Define se o arquivo precisa existir no momento da validacao.

    Returns:
        Path:
            Caminho absoluto do arquivo.

    Raises:
        TypeError:
            Caso caminho ou deve_existir possuam tipos invalidos.
        ArquivoDownloadNaoEncontradoError:
            Caso o arquivo obrigatorio nao exista.
    """
    # Valida o tipo do caminho informado
    if not isinstance(caminho, (str, Path)):
        raise TypeError(
            "O parametro 'caminho' deve ser uma string ou Path."
        )

    # Valida a opcao que controla a existencia do arquivo
    if not isinstance(deve_existir, bool):
        raise TypeError(
            "O parametro 'deve_existir' deve ser booleano."
        )

    # Converte o caminho para um caminho absoluto
    caminho_validado = Path(caminho).expanduser().resolve()

    # Interrompe a execucao quando o arquivo obrigatorio nao existe
    if deve_existir and not caminho_validado.is_file():
        raise ArquivoDownloadNaoEncontradoError(
            f"O arquivo de download nao foi encontrado: {caminho_validado}"
        )

    # Retorna o caminho normalizado
    return caminho_validado


def _arquivo_e_temporario(
    arquivo: Path,
    extensoes_temporarias: tuple[str, ...],
) -> bool:
    """
    Verifica se um arquivo utiliza uma extensao temporaria de download.

    Args:
        arquivo (Path):
            Arquivo que sera analisado.
        extensoes_temporarias (tuple[str, ...]):
            Extensoes que identificam downloads incompletos.

    Returns:
        bool:
            True quando o arquivo for temporario e False caso contrario.
    """
    # Compara o nome completo para contemplar arquivos como relatorio.xlsx.crdownload
    nome_arquivo = arquivo.name.lower()

    # Retorna True quando o nome termina com alguma extensao temporaria
    return any(
        nome_arquivo.endswith(extensao)
        for extensao in extensoes_temporarias
    )


# ----------------------------------------------------------------------------
# FUNCOES DE DIRETORIO E CONFIGURACAO
# ----------------------------------------------------------------------------

def preparar_diretorio_downloads(
    diretorio: str | Path = "downloads",
) -> Path:
    """
    Cria e retorna o diretorio utilizado para armazenar downloads.

    Args:
        diretorio (str | Path):
            Diretorio onde os arquivos baixados serao armazenados.

    Returns:
        Path:
            Caminho absoluto do diretorio criado ou localizado.

    Raises:
        TypeError:
            Caso diretorio nao seja uma string ou Path.
        DiretorioDownloadError:
            Caso exista um arquivo no caminho esperado do diretorio.
        OSError:
            Caso o sistema operacional nao consiga criar o diretorio.
    """
    # Verifica se o diretorio possui um tipo aceito
    if not isinstance(diretorio, (str, Path)):
        raise TypeError(
            "O parametro 'diretorio' deve ser uma string ou Path."
        )

    # Converte o diretorio para caminho absoluto
    caminho_diretorio = Path(diretorio).expanduser().resolve()

    # Impede o uso de um arquivo como diretorio
    if caminho_diretorio.exists() and not caminho_diretorio.is_dir():
        raise DiretorioDownloadError(
            f"O caminho informado nao e um diretorio: {caminho_diretorio}"
        )

    # Cria o diretorio e seus pais quando necessario
    caminho_diretorio.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Retorna o caminho preparado
    return caminho_diretorio


def configurar_downloads_chromium(
    driver: WebDriver,
    diretorio: str | Path,
    permitir: bool = True,
    eventos: bool = True,
) -> Path:
    """
    Configura o diretorio de downloads no Chrome ou Edge via DevTools.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        diretorio (str | Path):
            Diretorio utilizado pelo navegador para os downloads.
        permitir (bool):
            Define se downloads serao permitidos ou bloqueados.
        eventos (bool):
            Define se o navegador emitira eventos de progresso de download.

    Returns:
        Path:
            Caminho absoluto configurado no navegador.

    Raises:
        TypeError:
            Caso permitir ou eventos nao sejam booleanos.
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
        DownloadsError:
            Caso o navegador nao suporte o comando DevTools.
    """
    # Valida a sessao ativa
    _validar_driver(driver)

    # Verifica as opcoes booleanas
    if not isinstance(permitir, bool):
        raise TypeError(
            "O parametro 'permitir' deve ser booleano."
        )

    if not isinstance(eventos, bool):
        raise TypeError(
            "O parametro 'eventos' deve ser booleano."
        )

    # Prepara o diretorio antes de envia-lo ao navegador
    diretorio_preparado = preparar_diretorio_downloads(diretorio)

    # Define o comportamento conforme a permissao solicitada
    comportamento = "allow" if permitir else "deny"

    try:
        # Configura o comportamento de downloads no navegador Chromium
        driver.execute_cdp_cmd(
            "Browser.setDownloadBehavior",
            {
                "behavior": comportamento,
                "downloadPath": str(diretorio_preparado),
                "eventsEnabled": eventos,
            },
        )
    except WebDriverException as error:
        raise DownloadsError(
            "Nao foi possivel configurar os downloads pelo protocolo "
            "DevTools. Utilize Google Chrome ou Microsoft Edge."
        ) from error

    # Retorna o caminho aplicado ao navegador
    return diretorio_preparado


# ----------------------------------------------------------------------------
# FUNCOES DE CONSULTA
# ----------------------------------------------------------------------------

def listar_arquivos_download(
    diretorio: str | Path = "downloads",
    extensoes: tuple[str, ...] | list[str] | str | None = None,
    incluir_temporarios: bool = False,
    recursivo: bool = False,
) -> list[Path]:
    """
    Lista os arquivos existentes no diretorio de downloads.

    Args:
        diretorio (str | Path):
            Diretorio que sera consultado.
        extensoes (tuple | list | str | None):
            Extensoes utilizadas para filtrar os resultados.
        incluir_temporarios (bool):
            Define se arquivos temporarios serao incluidos.
        recursivo (bool):
            Define se subdiretorios tambem serao pesquisados.

    Returns:
        list[Path]:
            Arquivos ordenados do mais recente para o mais antigo.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        DiretorioDownloadError:
            Caso o diretorio informado nao exista.
    """
    # Valida as opcoes booleanas
    if not isinstance(incluir_temporarios, bool):
        raise TypeError(
            "O parametro 'incluir_temporarios' deve ser booleano."
        )

    if not isinstance(recursivo, bool):
        raise TypeError(
            "O parametro 'recursivo' deve ser booleano."
        )

    # Valida o tipo do diretorio sem cria-lo automaticamente
    if not isinstance(diretorio, (str, Path)):
        raise TypeError(
            "O parametro 'diretorio' deve ser uma string ou Path."
        )

    caminho_diretorio = Path(diretorio).expanduser().resolve()

    if not caminho_diretorio.is_dir():
        raise DiretorioDownloadError(
            f"O diretorio de downloads nao foi encontrado: {caminho_diretorio}"
        )

    # Normaliza os filtros de extensao
    extensoes_normalizadas = _normalizar_extensoes(extensoes)
    temporarias = _normalizar_extensoes(EXTENSOES_TEMPORARIAS_PADRAO)

    # Seleciona pesquisa simples ou recursiva
    arquivos_candidatos = (
        caminho_diretorio.rglob("*")
        if recursivo
        else caminho_diretorio.glob("*")
    )

    # Inicializa a lista de arquivos validos
    arquivos = []

    # Percorre os candidatos e aplica os filtros
    for arquivo in arquivos_candidatos:
        if not arquivo.is_file():
            continue

        if not incluir_temporarios and _arquivo_e_temporario(
            arquivo,
            temporarias,
        ):
            continue

        if (
            extensoes_normalizadas is not None
            and arquivo.suffix.lower() not in extensoes_normalizadas
        ):
            continue

        arquivos.append(arquivo)

    # Retorna os arquivos do mais recente para o mais antigo
    return sorted(
        arquivos,
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )


def obter_snapshot_downloads(
    diretorio: str | Path = "downloads",
    incluir_temporarios: bool = True,
) -> dict[str, tuple[int, int]]:
    """
    Registra um snapshot dos arquivos existentes no diretorio.

    Args:
        diretorio (str | Path):
            Diretorio que sera analisado.
        incluir_temporarios (bool):
            Define se arquivos temporarios serao registrados.

    Returns:
        dict[str, tuple[int, int]]:
            Mapeamento do caminho para tamanho e tempo de modificacao em ns.

    Raises:
        DiretorioDownloadError:
            Caso o diretorio informado nao exista.
        TypeError:
            Caso incluir_temporarios nao seja booleano.
    """
    # Lista os arquivos conforme a opcao solicitada
    arquivos = listar_arquivos_download(
        diretorio=diretorio,
        incluir_temporarios=incluir_temporarios,
    )

    # Cria um retrato contendo tamanho e modificacao de cada arquivo
    return {
        str(arquivo): (
            arquivo.stat().st_size,
            arquivo.stat().st_mtime_ns,
        )
        for arquivo in arquivos
    }


def identificar_novos_arquivos(
    diretorio: str | Path,
    snapshot_anterior: dict[str, tuple[int, int]] | set[str] | list[str],
    incluir_modificados: bool = False,
    incluir_temporarios: bool = True,
) -> list[Path]:
    """
    Identifica arquivos novos ou modificados apos um snapshot.

    Args:
        diretorio (str | Path):
            Diretorio que sera analisado.
        snapshot_anterior (dict | set | list):
            Estado dos arquivos existente antes do download.
        incluir_modificados (bool):
            Inclui arquivos existentes que tiveram tamanho ou data alterados.
        incluir_temporarios (bool):
            Define se arquivos temporarios serao considerados.

    Returns:
        list[Path]:
            Arquivos novos ou modificados, do mais recente para o mais antigo.

    Raises:
        TypeError:
            Caso o snapshot ou as opcoes possuam tipos invalidos.
    """
    # Valida as opcoes booleanas
    if not isinstance(incluir_modificados, bool):
        raise TypeError(
            "O parametro 'incluir_modificados' deve ser booleano."
        )

    if not isinstance(incluir_temporarios, bool):
        raise TypeError(
            "O parametro 'incluir_temporarios' deve ser booleano."
        )

    # Converte snapshots simplificados para um dicionario de caminhos
    if isinstance(snapshot_anterior, (set, list)):
        snapshot_normalizado = {
            str(Path(item).expanduser().resolve()): None
            for item in snapshot_anterior
        }
    elif isinstance(snapshot_anterior, dict):
        snapshot_normalizado = {
            str(Path(item).expanduser().resolve()): valor
            for item, valor in snapshot_anterior.items()
        }
    else:
        raise TypeError(
            "O parametro 'snapshot_anterior' deve ser dict, set ou lista."
        )

    # Lista o estado atual do diretorio
    arquivos_atuais = listar_arquivos_download(
        diretorio=diretorio,
        incluir_temporarios=incluir_temporarios,
    )

    # Inicializa a lista de arquivos detectados
    resultado = []

    # Compara cada arquivo atual com o snapshot anterior
    for arquivo in arquivos_atuais:
        caminho = str(arquivo.resolve())

        if caminho not in snapshot_normalizado:
            resultado.append(arquivo)
            continue

        if incluir_modificados and snapshot_normalizado[caminho] is not None:
            estado_atual = (
                arquivo.stat().st_size,
                arquivo.stat().st_mtime_ns,
            )

            if estado_atual != tuple(snapshot_normalizado[caminho]):
                resultado.append(arquivo)

    # Retorna os arquivos na ordem produzida pela listagem atual
    return resultado


def existe_download_em_andamento(
    diretorio: str | Path,
    extensoes_temporarias: tuple[str, ...] = EXTENSOES_TEMPORARIAS_PADRAO,
) -> bool:
    """
    Verifica se existem arquivos temporarios de download no diretorio.

    Args:
        diretorio (str | Path):
            Diretorio de downloads que sera consultado.
        extensoes_temporarias (tuple[str, ...]):
            Extensoes utilizadas para identificar downloads incompletos.

    Returns:
        bool:
            True quando existir download em andamento e False caso contrario.

    Raises:
        DiretorioDownloadError:
            Caso o diretorio informado nao exista.
        TypeError:
            Caso as extensoes informadas sejam invalidas.
    """
    # Normaliza as extensoes temporarias
    temporarias = _normalizar_extensoes(extensoes_temporarias)

    # Lista todos os arquivos, incluindo temporarios
    arquivos = listar_arquivos_download(
        diretorio=diretorio,
        incluir_temporarios=True,
    )

    # Retorna True quando algum arquivo possui extensao temporaria
    return any(
        _arquivo_e_temporario(arquivo, temporarias)
        for arquivo in arquivos
    )


def verificar_download_concluido(
    caminho_arquivo: str | Path,
    extensoes_temporarias: tuple[str, ...] = EXTENSOES_TEMPORARIAS_PADRAO,
    tamanho_minimo: int = 1,
) -> bool:
    """
    Verifica se um arquivo existe e aparenta estar concluido.

    Args:
        caminho_arquivo (str | Path):
            Caminho do arquivo que sera verificado.
        extensoes_temporarias (tuple[str, ...]):
            Extensoes que identificam downloads incompletos.
        tamanho_minimo (int):
            Tamanho minimo esperado em bytes.

    Returns:
        bool:
            True quando o arquivo estiver concluido e False caso contrario.

    Raises:
        TypeError:
            Caso algum parametro possua tipo invalido.
        ValueError:
            Caso tamanho_minimo seja negativo.
    """
    # Valida o caminho sem exigir que ele exista
    arquivo = _validar_caminho_arquivo(
        caminho_arquivo,
        deve_existir=False,
    )

    # Valida o tamanho minimo permitindo zero
    _validar_numero_positivo(
        tamanho_minimo,
        "tamanho_minimo",
        permitir_zero=True,
    )

    # Normaliza as extensoes temporarias
    temporarias = _normalizar_extensoes(extensoes_temporarias)

    # Retorna False quando o arquivo ainda nao existe
    if not arquivo.is_file():
        return False

    # Retorna False quando o arquivo utiliza extensao temporaria
    if _arquivo_e_temporario(arquivo, temporarias):
        return False

    # Retorna o resultado da validacao do tamanho minimo
    return arquivo.stat().st_size >= tamanho_minimo


# ----------------------------------------------------------------------------
# FUNCOES DE ESPERA
# ----------------------------------------------------------------------------

def aguardar_inicio_download(
    diretorio: str | Path,
    snapshot_anterior: dict[str, tuple[int, int]] | set[str] | list[str],
    timeout: int | float = 30,
    intervalo: int | float = 0.5,
) -> list[Path]:
    """
    Aguarda o surgimento de um novo arquivo no diretorio de downloads.

    Args:
        diretorio (str | Path):
            Diretorio monitorado.
        snapshot_anterior (dict | set | list):
            Estado do diretorio antes da acao de download.
        timeout (int | float):
            Tempo maximo para o download iniciar.
        intervalo (int | float):
            Intervalo entre as verificacoes.

    Returns:
        list[Path]:
            Arquivos novos, incluindo arquivos temporarios.

    Raises:
        DownloadTimeoutError:
            Caso nenhum arquivo novo apareca dentro do tempo limite.
    """
    # Valida os valores utilizados no loop de espera
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(intervalo, "intervalo")

    # Calcula o instante final da espera
    limite = time.monotonic() + timeout

    # Executa verificacoes ate o timeout
    while time.monotonic() < limite:
        novos = identificar_novos_arquivos(
            diretorio=diretorio,
            snapshot_anterior=snapshot_anterior,
            incluir_modificados=True,
            incluir_temporarios=True,
        )

        if novos:
            return novos

        time.sleep(intervalo)

    # Gera erro quando nenhum arquivo novo e identificado
    raise DownloadTimeoutError(
        "Nenhum download foi iniciado dentro do tempo limite."
    )


def aguardar_arquivo_estavel(
    caminho_arquivo: str | Path,
    timeout: int | float = 120,
    intervalo: int | float = 1,
    verificacoes_estaveis: int = 2,
    tamanho_minimo: int = 1,
) -> Path:
    """
    Aguarda o tamanho de um arquivo permanecer estavel.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera monitorado.
        timeout (int | float):
            Tempo maximo da espera.
        intervalo (int | float):
            Intervalo entre verificacoes de tamanho.
        verificacoes_estaveis (int):
            Quantidade de leituras consecutivas com o mesmo tamanho.
        tamanho_minimo (int):
            Tamanho minimo esperado em bytes.

    Returns:
        Path:
            Caminho absoluto do arquivo considerado estavel.

    Raises:
        DownloadTimeoutError:
            Caso o arquivo nao fique estavel dentro do tempo limite.
        TypeError:
            Caso verificacoes_estaveis nao seja inteiro.
    """
    # Normaliza o caminho sem exigir a existencia inicial
    arquivo = _validar_caminho_arquivo(
        caminho_arquivo,
        deve_existir=False,
    )

    # Valida os parametros numericos da espera
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(intervalo, "intervalo")
    _validar_numero_positivo(
        tamanho_minimo,
        "tamanho_minimo",
        permitir_zero=True,
    )

    if (
        isinstance(verificacoes_estaveis, bool)
        or not isinstance(verificacoes_estaveis, int)
    ):
        raise TypeError(
            "O parametro 'verificacoes_estaveis' deve ser inteiro."
        )

    if verificacoes_estaveis <= 0:
        raise ValueError(
            "O parametro 'verificacoes_estaveis' deve ser maior que zero."
        )

    # Inicializa os controles da estabilidade
    limite = time.monotonic() + timeout
    tamanho_anterior = None
    repeticoes = 0

    # Monitora o tamanho ate atingir a estabilidade solicitada
    while time.monotonic() < limite:
        if arquivo.is_file():
            tamanho_atual = arquivo.stat().st_size

            if tamanho_atual >= tamanho_minimo and tamanho_atual == tamanho_anterior:
                repeticoes += 1
            else:
                repeticoes = 0

            if repeticoes >= verificacoes_estaveis:
                return arquivo

            tamanho_anterior = tamanho_atual

        time.sleep(intervalo)

    # Gera erro quando o arquivo nao estabiliza
    raise DownloadTimeoutError(
        f"O arquivo nao ficou estavel dentro do tempo limite: {arquivo}"
    )


def aguardar_download(
    diretorio: str | Path,
    snapshot_anterior: dict[str, tuple[int, int]] | set[str] | list[str],
    timeout: int | float = 120,
    intervalo: int | float = 0.5,
    extensoes_esperadas: tuple[str, ...] | list[str] | str | None = None,
    tamanho_minimo: int = 1,
    aguardar_estabilidade: bool = True,
) -> Path:
    """
    Aguarda um novo download ser concluido e retorna seu caminho.

    Args:
        diretorio (str | Path):
            Diretorio monitorado.
        snapshot_anterior (dict | set | list):
            Estado dos arquivos antes da acao de download.
        timeout (int | float):
            Tempo maximo para conclusao.
        intervalo (int | float):
            Intervalo entre verificacoes.
        extensoes_esperadas (tuple | list | str | None):
            Extensoes permitidas para o arquivo concluido.
        tamanho_minimo (int):
            Tamanho minimo esperado em bytes.
        aguardar_estabilidade (bool):
            Define se o tamanho sera validado antes do retorno.

    Returns:
        Path:
            Caminho do arquivo baixado e concluido.

    Raises:
        DownloadTimeoutError:
            Caso nenhum download seja concluido dentro do tempo limite.
    """
    # Valida os parametros da espera
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(intervalo, "intervalo")
    _validar_numero_positivo(
        tamanho_minimo,
        "tamanho_minimo",
        permitir_zero=True,
    )

    if not isinstance(aguardar_estabilidade, bool):
        raise TypeError(
            "O parametro 'aguardar_estabilidade' deve ser booleano."
        )

    # Normaliza as extensoes esperadas
    extensoes_normalizadas = _normalizar_extensoes(extensoes_esperadas)
    temporarias = _normalizar_extensoes(EXTENSOES_TEMPORARIAS_PADRAO)

    # Define o instante limite da espera
    limite = time.monotonic() + timeout

    # Monitora o diretorio ate encontrar um arquivo concluido
    while time.monotonic() < limite:
        novos = identificar_novos_arquivos(
            diretorio=diretorio,
            snapshot_anterior=snapshot_anterior,
            incluir_modificados=True,
            incluir_temporarios=True,
        )

        # Avalia os arquivos mais recentes primeiro
        for arquivo in novos:
            if _arquivo_e_temporario(arquivo, temporarias):
                continue

            if (
                extensoes_normalizadas is not None
                and arquivo.suffix.lower() not in extensoes_normalizadas
            ):
                continue

            if arquivo.stat().st_size < tamanho_minimo:
                continue

            if aguardar_estabilidade:
                tempo_restante = max(0.1, limite - time.monotonic())
                return aguardar_arquivo_estavel(
                    arquivo,
                    timeout=tempo_restante,
                    intervalo=intervalo,
                    verificacoes_estaveis=2,
                    tamanho_minimo=tamanho_minimo,
                )

            return arquivo

        time.sleep(intervalo)

    # Gera erro quando nenhum arquivo concluido e encontrado
    raise DownloadTimeoutError(
        "Nenhum download foi concluido dentro do tempo limite."
    )


def aguardar_novos_downloads(
    diretorio: str | Path,
    snapshot_anterior: dict[str, tuple[int, int]] | set[str] | list[str],
    quantidade: int,
    timeout: int | float = 180,
    intervalo: int | float = 0.5,
    extensoes_esperadas: tuple[str, ...] | list[str] | str | None = None,
    tamanho_minimo: int = 1,
) -> list[Path]:
    """
    Aguarda uma quantidade exata de novos downloads concluidos.

    Args:
        diretorio (str | Path):
            Diretorio monitorado.
        snapshot_anterior (dict | set | list):
            Estado dos arquivos antes dos downloads.
        quantidade (int):
            Quantidade de arquivos concluidos esperada.
        timeout (int | float):
            Tempo maximo da espera.
        intervalo (int | float):
            Intervalo entre verificacoes.
        extensoes_esperadas (tuple | list | str | None):
            Extensoes permitidas nos resultados.
        tamanho_minimo (int):
            Tamanho minimo esperado para cada arquivo.

    Returns:
        list[Path]:
            Arquivos concluidos encontrados.

    Raises:
        DownloadTimeoutError:
            Caso a quantidade esperada nao seja atingida.
    """
    # Valida a quantidade esperada
    if isinstance(quantidade, bool) or not isinstance(quantidade, int):
        raise TypeError(
            "O parametro 'quantidade' deve ser inteiro."
        )

    if quantidade <= 0:
        raise ValueError(
            "O parametro 'quantidade' deve ser maior que zero."
        )

    # Valida os demais parametros numericos
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(intervalo, "intervalo")
    _validar_numero_positivo(
        tamanho_minimo,
        "tamanho_minimo",
        permitir_zero=True,
    )

    # Normaliza as extensoes utilizadas nos filtros
    extensoes_normalizadas = _normalizar_extensoes(extensoes_esperadas)
    temporarias = _normalizar_extensoes(EXTENSOES_TEMPORARIAS_PADRAO)
    limite = time.monotonic() + timeout

    # Monitora o diretorio ate atingir a quantidade esperada
    while time.monotonic() < limite:
        novos = identificar_novos_arquivos(
            diretorio=diretorio,
            snapshot_anterior=snapshot_anterior,
            incluir_modificados=True,
            incluir_temporarios=True,
        )

        concluidos = [
            arquivo
            for arquivo in novos
            if not _arquivo_e_temporario(arquivo, temporarias)
            and arquivo.stat().st_size >= tamanho_minimo
            and (
                extensoes_normalizadas is None
                or arquivo.suffix.lower() in extensoes_normalizadas
            )
        ]

        if len(concluidos) >= quantidade:
            return concluidos[:quantidade]

        time.sleep(intervalo)

    # Gera erro quando a quantidade nao e atingida
    raise DownloadTimeoutError(
        f"A quantidade de {quantidade} downloads nao foi concluida "
        "dentro do tempo limite."
    )


def aguardar_downloads_concluidos(
    diretorio: str | Path,
    timeout: int | float = 120,
    intervalo: int | float = 0.5,
) -> bool:
    """
    Aguarda todos os arquivos temporarios desaparecerem do diretorio.

    Args:
        diretorio (str | Path):
            Diretorio que sera monitorado.
        timeout (int | float):
            Tempo maximo para conclusao dos downloads.
        intervalo (int | float):
            Intervalo entre verificacoes.

    Returns:
        bool:
            True quando nao existirem downloads em andamento.

    Raises:
        DownloadTimeoutError:
            Caso ainda existam arquivos temporarios ao atingir o timeout.
    """
    # Valida os parametros de tempo
    _validar_numero_positivo(timeout, "timeout")
    _validar_numero_positivo(intervalo, "intervalo")

    # Define o instante limite da espera
    limite = time.monotonic() + timeout

    # Monitora o diretorio enquanto houver arquivos temporarios
    while time.monotonic() < limite:
        if not existe_download_em_andamento(diretorio):
            return True

        time.sleep(intervalo)

    # Gera erro quando ainda existem downloads em andamento
    raise DownloadTimeoutError(
        "Existem downloads em andamento apos o tempo limite."
    )


# ----------------------------------------------------------------------------
# FUNCOES DE LOCALIZACAO E VALIDACAO
# ----------------------------------------------------------------------------

def obter_arquivo_mais_recente(
    diretorio: str | Path,
    extensoes: tuple[str, ...] | list[str] | str | None = None,
    incluir_temporarios: bool = False,
) -> Path:
    """
    Retorna o arquivo modificado mais recentemente no diretorio.

    Args:
        diretorio (str | Path):
            Diretorio que sera consultado.
        extensoes (tuple | list | str | None):
            Extensoes utilizadas para filtrar os arquivos.
        incluir_temporarios (bool):
            Define se arquivos temporarios serao considerados.

    Returns:
        Path:
            Caminho absoluto do arquivo mais recente.

    Raises:
        ArquivoDownloadNaoEncontradoError:
            Caso nenhum arquivo correspondente seja encontrado.
    """
    # Lista os arquivos em ordem de modificacao decrescente
    arquivos = listar_arquivos_download(
        diretorio=diretorio,
        extensoes=extensoes,
        incluir_temporarios=incluir_temporarios,
    )

    # Interrompe a execucao quando nenhum arquivo e encontrado
    if not arquivos:
        raise ArquivoDownloadNaoEncontradoError(
            "Nenhum arquivo de download correspondente foi encontrado."
        )

    # Retorna o primeiro item porque a lista esta ordenada
    return arquivos[0]


def localizar_arquivo_download(
    diretorio: str | Path,
    nome: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    recursivo: bool = False,
) -> Path:
    """
    Localiza um arquivo de download pelo nome completo ou parcial.

    Args:
        diretorio (str | Path):
            Diretorio que sera pesquisado.
        nome (str):
            Nome completo ou trecho procurado.
        correspondencia_exata (bool):
            Define se o nome deve corresponder exatamente.
        considerar_maiusculas (bool):
            Define se a pesquisa diferencia maiusculas de minusculas.
        recursivo (bool):
            Define se subdiretorios tambem serao pesquisados.

    Returns:
        Path:
            Arquivo mais recente correspondente ao nome.

    Raises:
        ArquivoDownloadNaoEncontradoError:
            Caso nenhum arquivo correspondente seja encontrado.
        TypeError:
            Caso algum parametro possua tipo invalido.
    """
    # Valida o texto procurado
    if not isinstance(nome, str) or not nome.strip():
        raise TypeError(
            "O parametro 'nome' deve ser uma string nao vazia."
        )

    # Valida as opcoes booleanas
    for parametro, valor in {
        "correspondencia_exata": correspondencia_exata,
        "considerar_maiusculas": considerar_maiusculas,
        "recursivo": recursivo,
    }.items():
        if not isinstance(valor, bool):
            raise TypeError(
                f"O parametro '{parametro}' deve ser booleano."
            )

    # Lista os arquivos disponiveis
    arquivos = listar_arquivos_download(
        diretorio=diretorio,
        incluir_temporarios=False,
        recursivo=recursivo,
    )

    # Normaliza o nome procurado quando a pesquisa ignora maiusculas
    nome_procurado = nome.strip()

    if not considerar_maiusculas:
        nome_procurado = nome_procurado.lower()

    # Percorre os arquivos ja ordenados do mais recente para o mais antigo
    for arquivo in arquivos:
        nome_arquivo = arquivo.name

        if not considerar_maiusculas:
            nome_arquivo = nome_arquivo.lower()

        corresponde = (
            nome_arquivo == nome_procurado
            if correspondencia_exata
            else nome_procurado in nome_arquivo
        )

        if corresponde:
            return arquivo

    # Gera erro quando nenhum arquivo corresponde ao filtro
    raise ArquivoDownloadNaoEncontradoError(
        f"Nenhum arquivo correspondente a '{nome}' foi encontrado."
    )


def validar_arquivo_download(
    caminho_arquivo: str | Path,
    extensoes_permitidas: tuple[str, ...] | list[str] | str | None = None,
    tamanho_minimo: int = 1,
    permitir_vazio: bool = False,
) -> Path:
    """
    Valida existencia, extensao e tamanho de um arquivo baixado.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera validado.
        extensoes_permitidas (tuple | list | str | None):
            Extensoes aceitas para o arquivo.
        tamanho_minimo (int):
            Tamanho minimo esperado em bytes.
        permitir_vazio (bool):
            Define se um arquivo de tamanho zero sera aceito.

    Returns:
        Path:
            Caminho absoluto do arquivo validado.

    Raises:
        ArquivoDownloadNaoEncontradoError:
            Caso o arquivo nao exista.
        DownloadIncompletoError:
            Caso extensao ou tamanho nao atendam aos criterios.
    """
    # Localiza e normaliza o arquivo obrigatorio
    arquivo = _validar_caminho_arquivo(caminho_arquivo)

    # Valida a opcao de arquivo vazio
    if not isinstance(permitir_vazio, bool):
        raise TypeError(
            "O parametro 'permitir_vazio' deve ser booleano."
        )

    # Valida o tamanho minimo
    _validar_numero_positivo(
        tamanho_minimo,
        "tamanho_minimo",
        permitir_zero=True,
    )

    # Normaliza as extensoes permitidas
    extensoes = _normalizar_extensoes(extensoes_permitidas)

    # Verifica se a extensao do arquivo e aceita
    if extensoes is not None and arquivo.suffix.lower() not in extensoes:
        raise DownloadIncompletoError(
            f"A extensao '{arquivo.suffix}' nao e permitida. "
            f"Extensoes aceitas: {extensoes}"
        )

    # Recupera o tamanho atual do arquivo
    tamanho = arquivo.stat().st_size

    # Verifica se um arquivo vazio e permitido
    if tamanho == 0 and not permitir_vazio:
        raise DownloadIncompletoError(
            "O arquivo baixado esta vazio."
        )

    # Verifica o tamanho minimo solicitado
    if tamanho < tamanho_minimo:
        raise DownloadIncompletoError(
            f"O arquivo possui {tamanho} byte(s), abaixo do minimo de "
            f"{tamanho_minimo} byte(s)."
        )

    # Retorna o arquivo validado
    return arquivo


def calcular_hash_arquivo(
    caminho_arquivo: str | Path,
    algoritmo: AlgoritmoHash = "sha256",
    tamanho_bloco: int = 65536,
) -> str:
    """
    Calcula o hash de integridade de um arquivo baixado.

    Args:
        caminho_arquivo (str | Path):
            Arquivo utilizado no calculo.
        algoritmo (AlgoritmoHash):
            Algoritmo md5, sha1, sha256 ou sha512.
        tamanho_bloco (int):
            Quantidade de bytes lida a cada iteracao.

    Returns:
        str:
            Hash hexadecimal do arquivo.

    Raises:
        ValueError:
            Caso o algoritmo nao seja suportado ou o bloco seja invalido.
        ArquivoDownloadNaoEncontradoError:
            Caso o arquivo nao exista.
    """
    # Valida e localiza o arquivo
    arquivo = _validar_caminho_arquivo(caminho_arquivo)

    # Verifica se o algoritmo foi informado como string
    if not isinstance(algoritmo, str):
        raise TypeError(
            "O parametro 'algoritmo' deve ser uma string."
        )

    # Normaliza e valida o algoritmo
    algoritmo_normalizado = algoritmo.strip().lower()
    algoritmos_permitidos = {"md5", "sha1", "sha256", "sha512"}

    if algoritmo_normalizado not in algoritmos_permitidos:
        raise ValueError(
            f"Algoritmo nao suportado: '{algoritmo}'."
        )

    # Valida o tamanho do bloco de leitura
    if isinstance(tamanho_bloco, bool) or not isinstance(tamanho_bloco, int):
        raise TypeError(
            "O parametro 'tamanho_bloco' deve ser inteiro."
        )

    if tamanho_bloco <= 0:
        raise ValueError(
            "O parametro 'tamanho_bloco' deve ser maior que zero."
        )

    # Cria o objeto responsavel pelo calculo do hash
    hash_arquivo = hashlib.new(algoritmo_normalizado)

    # Le o arquivo em blocos para evitar consumo excessivo de memoria
    with arquivo.open("rb") as arquivo_binario:
        for bloco in iter(lambda: arquivo_binario.read(tamanho_bloco), b""):
            hash_arquivo.update(bloco)

    # Retorna o hash no formato hexadecimal
    return hash_arquivo.hexdigest()


# ----------------------------------------------------------------------------
# FUNCOES DE ORGANIZACAO DOS ARQUIVOS
# ----------------------------------------------------------------------------

def renomear_download(
    caminho_arquivo: str | Path,
    novo_nome: str,
    sobrescrever: bool = False,
    preservar_extensao: bool = True,
) -> Path:
    """
    Renomeia um arquivo baixado dentro do mesmo diretorio.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera renomeado.
        novo_nome (str):
            Novo nome do arquivo.
        sobrescrever (bool):
            Define se um arquivo existente podera ser substituido.
        preservar_extensao (bool):
            Adiciona a extensao original quando o novo nome nao possuir uma.

    Returns:
        Path:
            Novo caminho absoluto do arquivo.

    Raises:
        ArquivoDownloadExistenteError:
            Caso o destino exista e sobrescrever seja False.
        TypeError:
            Caso algum parametro possua tipo invalido.
    """
    # Valida e localiza o arquivo original
    arquivo = _validar_caminho_arquivo(caminho_arquivo)

    # Valida o novo nome
    if not isinstance(novo_nome, str) or not novo_nome.strip():
        raise TypeError(
            "O parametro 'novo_nome' deve ser uma string nao vazia."
        )

    # Impede que novo_nome contenha diretorios
    if Path(novo_nome).name != novo_nome:
        raise ValueError(
            "O parametro 'novo_nome' deve conter apenas o nome do arquivo."
        )

    # Valida as opcoes booleanas
    if not isinstance(sobrescrever, bool):
        raise TypeError(
            "O parametro 'sobrescrever' deve ser booleano."
        )

    if not isinstance(preservar_extensao, bool):
        raise TypeError(
            "O parametro 'preservar_extensao' deve ser booleano."
        )

    # Monta o nome final e preserva a extensao quando solicitado
    nome_final = novo_nome.strip()

    if preservar_extensao and not Path(nome_final).suffix:
        nome_final = f"{nome_final}{arquivo.suffix}"

    destino = arquivo.parent / nome_final

    # Verifica se o destino ja existe
    if destino.exists() and not sobrescrever:
        raise ArquivoDownloadExistenteError(
            f"Ja existe um arquivo no destino: {destino}"
        )

    # Substitui o destino existente quando permitido
    if destino.exists() and sobrescrever:
        destino.unlink()

    # Renomeia o arquivo e retorna o novo caminho
    return arquivo.rename(destino)


def mover_download(
    caminho_arquivo: str | Path,
    diretorio_destino: str | Path,
    novo_nome: str | None = None,
    sobrescrever: bool = False,
) -> Path:
    """
    Move um download para outro diretorio.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera movido.
        diretorio_destino (str | Path):
            Diretorio que recebera o arquivo.
        novo_nome (str | None):
            Nome opcional aplicado ao arquivo movido.
        sobrescrever (bool):
            Define se um arquivo existente podera ser substituido.

    Returns:
        Path:
            Caminho final do arquivo movido.

    Raises:
        ArquivoDownloadExistenteError:
            Caso o destino exista e sobrescrever seja False.
    """
    # Valida o arquivo original e prepara o diretorio de destino
    arquivo = _validar_caminho_arquivo(caminho_arquivo)
    destino_diretorio = preparar_diretorio_downloads(diretorio_destino)

    # Valida a opcao de sobrescrita
    if not isinstance(sobrescrever, bool):
        raise TypeError(
            "O parametro 'sobrescrever' deve ser booleano."
        )

    # Define o nome final do arquivo
    nome_final = arquivo.name

    if novo_nome is not None:
        if not isinstance(novo_nome, str) or not novo_nome.strip():
            raise TypeError(
                "O parametro 'novo_nome' deve ser string nao vazia ou None."
            )

        if Path(novo_nome).name != novo_nome:
            raise ValueError(
                "O parametro 'novo_nome' deve conter apenas o nome do arquivo."
            )

        nome_final = novo_nome.strip()

        if not Path(nome_final).suffix:
            nome_final = f"{nome_final}{arquivo.suffix}"

    destino = destino_diretorio / nome_final

    # Verifica e trata a existencia do destino
    if destino.exists() and not sobrescrever:
        raise ArquivoDownloadExistenteError(
            f"Ja existe um arquivo no destino: {destino}"
        )

    if destino.exists() and sobrescrever:
        destino.unlink()

    # Move o arquivo e retorna o caminho final
    return Path(shutil.move(str(arquivo), str(destino))).resolve()


def copiar_download(
    caminho_arquivo: str | Path,
    diretorio_destino: str | Path,
    novo_nome: str | None = None,
    sobrescrever: bool = False,
) -> Path:
    """
    Copia um download para outro diretorio preservando o original.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera copiado.
        diretorio_destino (str | Path):
            Diretorio que recebera a copia.
        novo_nome (str | None):
            Nome opcional aplicado a copia.
        sobrescrever (bool):
            Define se uma copia existente podera ser substituida.

    Returns:
        Path:
            Caminho final da copia criada.

    Raises:
        ArquivoDownloadExistenteError:
            Caso o destino exista e sobrescrever seja False.
    """
    # Valida o arquivo original e prepara o diretorio de destino
    arquivo = _validar_caminho_arquivo(caminho_arquivo)
    destino_diretorio = preparar_diretorio_downloads(diretorio_destino)

    # Valida a opcao de sobrescrita
    if not isinstance(sobrescrever, bool):
        raise TypeError(
            "O parametro 'sobrescrever' deve ser booleano."
        )

    # Define o nome da copia
    nome_final = arquivo.name

    if novo_nome is not None:
        if not isinstance(novo_nome, str) or not novo_nome.strip():
            raise TypeError(
                "O parametro 'novo_nome' deve ser string nao vazia ou None."
            )

        if Path(novo_nome).name != novo_nome:
            raise ValueError(
                "O parametro 'novo_nome' deve conter apenas o nome do arquivo."
            )

        nome_final = novo_nome.strip()

        if not Path(nome_final).suffix:
            nome_final = f"{nome_final}{arquivo.suffix}"

    destino = destino_diretorio / nome_final

    # Verifica e trata a existencia do destino
    if destino.exists() and not sobrescrever:
        raise ArquivoDownloadExistenteError(
            f"Ja existe um arquivo no destino: {destino}"
        )

    if destino.exists() and sobrescrever:
        destino.unlink()

    # Copia metadados e conteudo do arquivo
    return Path(shutil.copy2(str(arquivo), str(destino))).resolve()


def remover_download(
    caminho_arquivo: str | Path,
    ignorar_inexistente: bool = False,
) -> bool:
    """
    Remove um arquivo de download do sistema de arquivos.

    Args:
        caminho_arquivo (str | Path):
            Arquivo que sera removido.
        ignorar_inexistente (bool):
            Define se a ausencia do arquivo sera ignorada.

    Returns:
        bool:
            True quando o arquivo for removido e False quando for ignorado.

    Raises:
        ArquivoDownloadNaoEncontradoError:
            Caso o arquivo nao exista e ignorar_inexistente seja False.
    """
    # Valida a opcao de ausencia do arquivo
    if not isinstance(ignorar_inexistente, bool):
        raise TypeError(
            "O parametro 'ignorar_inexistente' deve ser booleano."
        )

    # Normaliza o caminho sem exigir existencia
    arquivo = _validar_caminho_arquivo(
        caminho_arquivo,
        deve_existir=False,
    )

    # Trata arquivos inexistentes conforme a opcao informada
    if not arquivo.is_file():
        if ignorar_inexistente:
            return False

        raise ArquivoDownloadNaoEncontradoError(
            f"O arquivo de download nao foi encontrado: {arquivo}"
        )

    # Remove o arquivo existente
    arquivo.unlink()

    # Retorna True para indicar a remocao
    return True


def limpar_diretorio_downloads(
    diretorio: str | Path,
    extensoes: tuple[str, ...] | list[str] | str | None = None,
    incluir_temporarios: bool = True,
    recursivo: bool = False,
) -> list[Path]:
    """
    Remove arquivos selecionados do diretorio de downloads.

    Args:
        diretorio (str | Path):
            Diretorio que sera limpo.
        extensoes (tuple | list | str | None):
            Extensoes que serao removidas. None remove todas.
        incluir_temporarios (bool):
            Define se arquivos temporarios serao removidos.
        recursivo (bool):
            Define se arquivos em subdiretorios serao removidos.

    Returns:
        list[Path]:
            Caminhos dos arquivos removidos.

    Raises:
        DiretorioDownloadError:
            Caso o diretorio informado nao exista.
    """
    # Lista os arquivos que atendem aos filtros informados
    arquivos = listar_arquivos_download(
        diretorio=diretorio,
        extensoes=extensoes,
        incluir_temporarios=incluir_temporarios,
        recursivo=recursivo,
    )

    # Inicializa a relacao de arquivos removidos
    removidos = []

    # Remove cada arquivo encontrado
    for arquivo in arquivos:
        arquivo.unlink()
        removidos.append(arquivo)

    # Retorna os caminhos removidos
    return removidos


# ----------------------------------------------------------------------------
# FUNCOES INTEGRADAS COM SELENIUM
# ----------------------------------------------------------------------------

def clicar_e_aguardar_download(
    driver: WebDriver,
    localizador: Localizador,
    diretorio: str | Path,
    timeout_clique: int | float = 20,
    timeout_download: int | float = 120,
    intervalo: int | float = 0.5,
    extensoes_esperadas: tuple[str, ...] | list[str] | str | None = None,
    tamanho_minimo: int = 1,
) -> Path:
    """
    Clica em um elemento e aguarda a conclusao do novo download.

    Args:
        driver (WebDriver):
            Instancia ativa do Selenium WebDriver.
        localizador (Localizador):
            Localizador do botao ou link de download.
        diretorio (str | Path):
            Diretorio monitorado.
        timeout_clique (int | float):
            Tempo maximo para o elemento ficar clicavel.
        timeout_download (int | float):
            Tempo maximo para o download ser concluido.
        intervalo (int | float):
            Intervalo entre verificacoes do diretorio.
        extensoes_esperadas (tuple | list | str | None):
            Extensoes permitidas para o arquivo baixado.
        tamanho_minimo (int):
            Tamanho minimo esperado em bytes.

    Returns:
        Path:
            Caminho do arquivo baixado e concluido.

    Raises:
        DownloadTimeoutError:
            Caso o elemento nao fique clicavel ou o download nao termine.
        DriverInativoError:
            Caso o WebDriver nao possua uma sessao ativa.
    """
    # Valida a sessao, o localizador e os tempos
    _validar_driver(driver)
    _validar_localizador(localizador)
    _validar_numero_positivo(timeout_clique, "timeout_clique")
    _validar_numero_positivo(timeout_download, "timeout_download")
    _validar_numero_positivo(intervalo, "intervalo")

    # Prepara o diretorio e registra seu estado antes do clique
    diretorio_preparado = preparar_diretorio_downloads(diretorio)
    snapshot = obter_snapshot_downloads(
        diretorio_preparado,
        incluir_temporarios=True,
    )

    try:
        # Aguarda o elemento ficar clicavel
        elemento: WebElement = WebDriverWait(
            driver,
            timeout_clique,
        ).until(
            EC.element_to_be_clickable(localizador)
        )

        # Executa o clique que inicia o download
        elemento.click()

    except TimeoutException as error:
        raise DownloadTimeoutError(
            f"O elemento {localizador} nao ficou clicavel dentro do tempo limite."
        ) from error
    except WebDriverException as error:
        raise DownloadsError(
            f"Nao foi possivel clicar no elemento {localizador}."
        ) from error

    # Aguarda e retorna o arquivo concluido
    return aguardar_download(
        diretorio=diretorio_preparado,
        snapshot_anterior=snapshot,
        timeout=timeout_download,
        intervalo=intervalo,
        extensoes_esperadas=extensoes_esperadas,
        tamanho_minimo=tamanho_minimo,
        aguardar_estabilidade=True,
    )
