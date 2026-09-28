"""
============================================================
Modulo: SELENIUM / excecoes.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Modulo central de excecoes personalizadas utilizado pela
    biblioteca de funcoes auxiliares para automacoes Selenium.

    O modulo organiza as excecoes por area de responsabilidade,
    permitindo tratamento generico ou especifico dos erros.

Grupos de excecoes:
    - Erros gerais da biblioteca
    - Erros de WebDriver e navegadores
    - Erros de elementos e localizadores
    - Erros de esperas explicitas
    - Erros de capturas e evidencias
    - Erros de downloads e arquivos
    - Erros de navegacao, janelas, frames e alertas
    - Erros de JavaScript e execucao de scripts

Dependencias:
    - Nenhuma dependencia externa

Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Centralizacao das excecoes dos modulos Selenium.
        - Inclusao de contexto estruturado nos erros.
        - Inclusao de conversao das excecoes para dicionario.
        - Inclusao de hierarquia por area de responsabilidade.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa Any para permitir valores genericos no contexto das excecoes
from typing import Any


# ----------------------------------------------------------------------------
# EXCECAO BASE DA BIBLIOTECA
# ----------------------------------------------------------------------------

class AutomacaoError(Exception):
    """
    Excecao base de toda a biblioteca de automacao.

    Todas as excecoes personalizadas do projeto herdam direta ou
    indiretamente desta classe. Isso permite capturar qualquer erro da
    biblioteca utilizando apenas except AutomacaoError.

    Args:
        mensagem (str):
            Descricao principal do erro ocorrido.
        codigo (str | None):
            Codigo opcional utilizado para identificar o tipo do erro.
        contexto (dict[str, Any] | None):
            Informacoes adicionais relacionadas ao erro.

    Attributes:
        mensagem (str):
            Mensagem principal normalizada.
        codigo (str | None):
            Codigo opcional do erro.
        contexto (dict[str, Any]):
            Dicionario contendo detalhes adicionais.

    Raises:
        TypeError:
            Caso mensagem, codigo ou contexto possuam tipos invalidos.
        ValueError:
            Caso a mensagem informada esteja vazia.
    """

    # Codigo padrao utilizado quando a subclasse nao define outro valor
    codigo_padrao = "AUTOMACAO_ERROR"

    def __init__(
        self,
        mensagem: str,
        codigo: str | None = None,
        contexto: dict[str, Any] | None = None,
    ) -> None:
        """
        Inicializa uma excecao personalizada da biblioteca.

        Args:
            mensagem (str):
                Descricao principal do erro ocorrido.
            codigo (str | None):
                Codigo opcional que substitui o codigo padrao da classe.
            contexto (dict[str, Any] | None):
                Informacoes adicionais relacionadas ao erro.

        Returns:
            None:
                O metodo apenas inicializa a excecao.

        Raises:
            TypeError:
                Caso algum parametro possua tipo invalido.
            ValueError:
                Caso a mensagem esteja vazia.
        """
        # Verifica se a mensagem informada e uma string
        if not isinstance(mensagem, str):
            raise TypeError(
                "O parametro 'mensagem' deve ser uma string."
            )

        # Remove espacos desnecessarios da mensagem
        mensagem_normalizada = mensagem.strip()

        # Interrompe a inicializacao quando a mensagem esta vazia
        if not mensagem_normalizada:
            raise ValueError(
                "O parametro 'mensagem' nao pode estar vazio."
            )

        # Valida o codigo quando ele e informado manualmente
        if codigo is not None and not isinstance(codigo, str):
            raise TypeError(
                "O parametro 'codigo' deve ser uma string ou None."
            )

        # Valida o contexto quando ele e informado
        if contexto is not None and not isinstance(contexto, dict):
            raise TypeError(
                "O parametro 'contexto' deve ser um dicionario ou None."
            )

        # Armazena a mensagem normalizada na instancia
        self.mensagem = mensagem_normalizada

        # Utiliza o codigo recebido ou o codigo padrao da subclasse
        self.codigo = (
            codigo.strip()
            if isinstance(codigo, str) and codigo.strip()
            else self.codigo_padrao
        )

        # Cria uma copia do contexto para evitar alteracoes externas
        self.contexto = dict(contexto or {})

        # Inicializa a classe Exception com a mensagem principal
        super().__init__(self.mensagem)

    def adicionar_contexto(
        self,
        chave: str,
        valor: Any,
    ) -> "AutomacaoError":
        """
        Adiciona uma informacao ao contexto da excecao.

        Args:
            chave (str):
                Nome utilizado para identificar a informacao.
            valor (Any):
                Valor que sera armazenado no contexto.

        Returns:
            AutomacaoError:
                A propria instancia, permitindo chamadas encadeadas.

        Raises:
            TypeError:
                Caso a chave nao seja uma string.
            ValueError:
                Caso a chave esteja vazia.
        """
        # Verifica se a chave informada e uma string
        if not isinstance(chave, str):
            raise TypeError(
                "O parametro 'chave' deve ser uma string."
            )

        # Remove espacos das extremidades da chave
        chave_normalizada = chave.strip()

        # Interrompe a execucao quando a chave esta vazia
        if not chave_normalizada:
            raise ValueError(
                "O parametro 'chave' nao pode estar vazio."
            )

        # Registra ou atualiza a informacao no contexto
        self.contexto[chave_normalizada] = valor

        # Retorna a propria excecao para permitir encadeamento
        return self

    def para_dicionario(self) -> dict[str, Any]:
        """
        Converte a excecao para uma estrutura serializavel.

        Returns:
            dict[str, Any]:
                Dicionario com tipo, codigo, mensagem e contexto do erro.
        """
        # Retorna uma copia dos dados associados a excecao
        return {
            "tipo": self.__class__.__name__,
            "codigo": self.codigo,
            "mensagem": self.mensagem,
            "contexto": dict(self.contexto),
        }

    def __str__(self) -> str:
        """
        Retorna a representacao textual da excecao.

        Returns:
            str:
                Codigo e mensagem, seguidos do contexto quando existente.
        """
        # Monta a representacao principal do erro
        texto = f"[{self.codigo}] {self.mensagem}"

        # Adiciona o contexto somente quando existem informacoes extras
        if self.contexto:
            texto = f"{texto} | Contexto: {self.contexto}"

        # Retorna a representacao final
        return texto


# Mantem um nome alternativo para compatibilidade com outros modulos
class RPAError(AutomacaoError):
    """Excecao base alternativa para erros gerais de RPA."""

    codigo_padrao = "RPA_ERROR"


# ----------------------------------------------------------------------------
# ERROS GERAIS DE VALIDACAO E CONFIGURACAO
# ----------------------------------------------------------------------------

class ValidacaoError(AutomacaoError):
    """Erro base para parametros, valores ou estruturas invalidas."""

    codigo_padrao = "VALIDACAO_ERROR"


class ParametroInvalidoError(ValidacaoError):
    """Um parametro da funcao possui tipo ou valor invalido."""

    codigo_padrao = "PARAMETRO_INVALIDO"


class ConfiguracaoError(AutomacaoError):
    """Erro base para configuracoes ausentes ou invalidas."""

    codigo_padrao = "CONFIGURACAO_ERROR"


class CaminhoInvalidoError(ValidacaoError):
    """Um caminho informado nao existe ou possui tipo incorreto."""

    codigo_padrao = "CAMINHO_INVALIDO"


class ArquivoNaoEncontradoError(CaminhoInvalidoError):
    """Um arquivo obrigatorio nao foi encontrado."""

    codigo_padrao = "ARQUIVO_NAO_ENCONTRADO"


class DiretorioNaoEncontradoError(CaminhoInvalidoError):
    """Um diretorio obrigatorio nao foi encontrado."""

    codigo_padrao = "DIRETORIO_NAO_ENCONTRADO"


# ----------------------------------------------------------------------------
# ERROS DE SELENIUM E WEBDRIVER
# ----------------------------------------------------------------------------

class SeleniumError(AutomacaoError):
    """Erro base para operacoes executadas com Selenium."""

    codigo_padrao = "SELENIUM_ERROR"


class DriverError(SeleniumError):
    """Erro base para operacoes relacionadas ao WebDriver."""

    codigo_padrao = "DRIVER_ERROR"


class DriverInativoError(DriverError):
    """O WebDriver nao possui uma sessao ativa ou valida."""

    codigo_padrao = "DRIVER_INATIVO"


class DriverInicializacaoError(DriverError):
    """O navegador ou WebDriver nao pode ser inicializado."""

    codigo_padrao = "DRIVER_INICIALIZACAO"


class DriverEncerramentoError(DriverError):
    """O WebDriver nao pode ser encerrado corretamente."""

    codigo_padrao = "DRIVER_ENCERRAMENTO"


class CaminhoDriverError(DriverError, CaminhoInvalidoError):
    """O caminho do driver, navegador ou perfil e invalido."""

    codigo_padrao = "CAMINHO_DRIVER_INVALIDO"


class NavegadorNaoSuportadoError(DriverError):
    """O navegador informado nao e suportado pela biblioteca."""

    codigo_padrao = "NAVEGADOR_NAO_SUPORTADO"


class SessaoDriverError(DriverError):
    """A sessao do WebDriver foi perdida, encerrada ou corrompida."""

    codigo_padrao = "SESSAO_DRIVER_ERROR"


# ----------------------------------------------------------------------------
# ERROS DE ELEMENTOS E LOCALIZADORES
# ----------------------------------------------------------------------------

class ElementosError(SeleniumError):
    """Erro base para operacoes relacionadas a elementos web."""

    codigo_padrao = "ELEMENTOS_ERROR"


class LocalizadorInvalidoError(ElementosError, ValidacaoError):
    """O localizador Selenium nao possui uma estrutura valida."""

    codigo_padrao = "LOCALIZADOR_INVALIDO"


class ElementoNaoEncontradoError(ElementosError):
    """O elemento solicitado nao foi encontrado na pagina."""

    codigo_padrao = "ELEMENTO_NAO_ENCONTRADO"


class ElementoNaoVisivelError(ElementosError):
    """O elemento existe no DOM, mas nao esta visivel."""

    codigo_padrao = "ELEMENTO_NAO_VISIVEL"


class ElementoNaoClicavelError(ElementosError):
    """O elemento nao ficou disponivel para clique."""

    codigo_padrao = "ELEMENTO_NAO_CLICAVEL"


class ElementoDesabilitadoError(ElementosError):
    """O elemento esta desabilitado para interacao."""

    codigo_padrao = "ELEMENTO_DESABILITADO"


class ElementoObsoletoError(ElementosError):
    """O elemento deixou de estar associado ao DOM atual."""

    codigo_padrao = "ELEMENTO_OBSOLETO"


class AtributoElementoError(ElementosError):
    """Um atributo ou propriedade do elemento nao pode ser obtido."""

    codigo_padrao = "ATRIBUTO_ELEMENTO_ERROR"


class InteracaoElementoError(ElementosError):
    """Uma interacao com o elemento nao pode ser concluida."""

    codigo_padrao = "INTERACAO_ELEMENTO_ERROR"


# ----------------------------------------------------------------------------
# ERROS DE ESPERAS EXPLICITAS
# ----------------------------------------------------------------------------

class EsperaError(SeleniumError):
    """Erro base para operacoes de espera e sincronizacao."""

    codigo_padrao = "ESPERA_ERROR"


class CondicaoInvalidaError(EsperaError, ValidacaoError):
    """A condicao personalizada informada nao e valida."""

    codigo_padrao = "CONDICAO_INVALIDA"


class TempoEsperaExcedidoError(EsperaError):
    """A condicao esperada nao foi atendida no tempo limite."""

    codigo_padrao = "TEMPO_ESPERA_EXCEDIDO"


class CarregamentoPaginaError(EsperaError):
    """A pagina nao concluiu o carregamento no tempo esperado."""

    codigo_padrao = "CARREGAMENTO_PAGINA_ERROR"


# ----------------------------------------------------------------------------
# ERROS DE CAPTURAS E EVIDENCIAS
# ----------------------------------------------------------------------------

class CapturasError(SeleniumError):
    """Erro base para capturas de tela e evidencias."""

    codigo_padrao = "CAPTURAS_ERROR"


class CaminhoCapturaError(CapturasError, CaminhoInvalidoError):
    """O caminho informado para uma captura nao e valido."""

    codigo_padrao = "CAMINHO_CAPTURA_INVALIDO"


class CapturaNaoRealizadaError(CapturasError):
    """A captura de tela ou elemento nao pode ser realizada."""

    codigo_padrao = "CAPTURA_NAO_REALIZADA"


class CapturaTelaCompletaError(CapturasError):
    """A captura da pagina completa nao pode ser realizada."""

    codigo_padrao = "CAPTURA_TELA_COMPLETA_ERROR"


class EvidenciaNaoEncontradaError(CapturasError):
    """Nenhuma evidencia correspondente foi encontrada."""

    codigo_padrao = "EVIDENCIA_NAO_ENCONTRADA"


class EvidenciaNaoSalvaError(CapturasError):
    """Uma evidencia nao pode ser gravada no sistema de arquivos."""

    codigo_padrao = "EVIDENCIA_NAO_SALVA"


class LogNavegadorError(CapturasError):
    """Os logs do navegador nao puderam ser obtidos ou salvos."""

    codigo_padrao = "LOG_NAVEGADOR_ERROR"


# ----------------------------------------------------------------------------
# ERROS DE DOWNLOADS
# ----------------------------------------------------------------------------

class DownloadsError(SeleniumError):
    """Erro base para operacoes relacionadas a downloads."""

    codigo_padrao = "DOWNLOADS_ERROR"


class DiretorioDownloadError(DownloadsError, CaminhoInvalidoError):
    """O diretorio de downloads nao existe ou nao e valido."""

    codigo_padrao = "DIRETORIO_DOWNLOAD_INVALIDO"


class ArquivoDownloadNaoEncontradoError(DownloadsError):
    """O arquivo baixado solicitado nao foi encontrado."""

    codigo_padrao = "ARQUIVO_DOWNLOAD_NAO_ENCONTRADO"


class ArquivoDownloadExistenteError(DownloadsError):
    """Ja existe um arquivo no caminho de destino informado."""

    codigo_padrao = "ARQUIVO_DOWNLOAD_EXISTENTE"


class DownloadTimeoutError(DownloadsError):
    """O download nao iniciou ou nao terminou no tempo definido."""

    codigo_padrao = "DOWNLOAD_TIMEOUT"


class DownloadIncompletoError(DownloadsError):
    """O download esta incompleto ou nao atende aos criterios definidos."""

    codigo_padrao = "DOWNLOAD_INCOMPLETO"


class DownloadNaoIniciadoError(DownloadsError):
    """Nenhum novo arquivo foi identificado apos a acao de download."""

    codigo_padrao = "DOWNLOAD_NAO_INICIADO"


class DownloadEmAndamentoError(DownloadsError):
    """Ainda existem arquivos temporarios de download no diretorio."""

    codigo_padrao = "DOWNLOAD_EM_ANDAMENTO"


class ExtensaoDownloadInvalidaError(DownloadsError, ValidacaoError):
    """A extensao do arquivo baixado nao e permitida."""

    codigo_padrao = "EXTENSAO_DOWNLOAD_INVALIDA"


class TamanhoDownloadInvalidoError(DownloadsError, ValidacaoError):
    """O tamanho do arquivo baixado nao atende ao limite definido."""

    codigo_padrao = "TAMANHO_DOWNLOAD_INVALIDO"


class HashArquivoError(DownloadsError):
    """O hash de integridade do arquivo nao pode ser calculado."""

    codigo_padrao = "HASH_ARQUIVO_ERROR"


# ----------------------------------------------------------------------------
# ERROS DE NAVEGACAO, JANELAS, FRAMES E ALERTAS
# ----------------------------------------------------------------------------

class NavegacaoError(SeleniumError):
    """Erro base para navegacao entre paginas e contextos."""

    codigo_padrao = "NAVEGACAO_ERROR"


class UrlInvalidaError(NavegacaoError, ValidacaoError):
    """A URL informada nao possui formato valido."""

    codigo_padrao = "URL_INVALIDA"


class JanelaError(NavegacaoError):
    """Erro base para operacoes com janelas e abas."""

    codigo_padrao = "JANELA_ERROR"


class JanelaNaoEncontradaError(JanelaError):
    """A janela ou aba solicitada nao foi encontrada."""

    codigo_padrao = "JANELA_NAO_ENCONTRADA"


class FrameError(NavegacaoError):
    """Erro base para operacoes relacionadas a frames e iframes."""

    codigo_padrao = "FRAME_ERROR"


class FrameNaoEncontradoError(FrameError):
    """O frame solicitado nao foi encontrado ou nao ficou disponivel."""

    codigo_padrao = "FRAME_NAO_ENCONTRADO"


class AlertaError(NavegacaoError):
    """Erro base para operacoes relacionadas a alertas nativos."""

    codigo_padrao = "ALERTA_ERROR"


class AlertaNaoEncontradoError(AlertaError):
    """Nenhum alerta nativo foi encontrado no tempo definido."""

    codigo_padrao = "ALERTA_NAO_ENCONTRADO"


# ----------------------------------------------------------------------------
# ERROS DE JAVASCRIPT
# ----------------------------------------------------------------------------

class JavaScriptError(SeleniumError):
    """Erro base para execucao de JavaScript no navegador."""

    codigo_padrao = "JAVASCRIPT_ERROR"


class ScriptInvalidoError(JavaScriptError, ValidacaoError):
    """O codigo JavaScript informado esta vazio ou e invalido."""

    codigo_padrao = "SCRIPT_INVALIDO"


class ExecucaoScriptError(JavaScriptError):
    """O navegador nao conseguiu executar o codigo JavaScript."""

    codigo_padrao = "EXECUCAO_SCRIPT_ERROR"


# ----------------------------------------------------------------------------
# FUNCOES AUXILIARES DO MODULO
# ----------------------------------------------------------------------------

def obter_hierarquia_excecao(
    excecao: BaseException | type[BaseException],
) -> list[str]:
    """
    Retorna a hierarquia de classes de uma excecao.

    Args:
        excecao (BaseException | type[BaseException]):
            Instancia ou classe cuja hierarquia sera consultada.

    Returns:
        list[str]:
            Nomes das classes, da mais especifica ate BaseException.

    Raises:
        TypeError:
            Caso o valor informado nao seja uma excecao ou classe de excecao.
    """
    # Identifica a classe quando uma instancia foi informada
    if isinstance(excecao, BaseException):
        classe = excecao.__class__
    elif isinstance(excecao, type) and issubclass(excecao, BaseException):
        classe = excecao
    else:
        raise TypeError(
            "O parametro 'excecao' deve ser uma excecao ou classe de excecao."
        )

    # Retorna os nomes das classes presentes no MRO da excecao
    return [
        item.__name__
        for item in classe.mro()
        if issubclass(item, BaseException)
    ]


def excecao_para_dicionario(
    excecao: BaseException,
) -> dict[str, Any]:
    """
    Converte uma excecao personalizada ou nativa para dicionario.

    Args:
        excecao (BaseException):
            Excecao que sera convertida.

    Returns:
        dict[str, Any]:
            Informacoes estruturadas sobre a excecao.

    Raises:
        TypeError:
            Caso o objeto informado nao seja uma excecao.
    """
    # Verifica se o objeto informado e uma excecao
    if not isinstance(excecao, BaseException):
        raise TypeError(
            "O parametro 'excecao' deve ser uma instancia de BaseException."
        )

    # Utiliza a conversao completa das excecoes da biblioteca
    if isinstance(excecao, AutomacaoError):
        return excecao.para_dicionario()

    # Cria uma estrutura minima para excecoes externas ou nativas
    return {
        "tipo": excecao.__class__.__name__,
        "codigo": None,
        "mensagem": str(excecao),
        "contexto": {},
    }
