"""
============================================================
Modulo: OUTLOOK / conexao.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes auxiliares para conectar ao Outlook classico, obter a
    aplicacao COM, acessar o Namespace MAPI, inicializar a sessao,
    consultar informacoes e liberar referencias utilizadas por RPAs.
Funcoes disponiveis:
    - verificar_pywin32_disponivel()
    - importar_win32com()
    - obter_outlook_ativo()
    - criar_aplicacao_outlook()
    - obter_aplicacao_outlook()
    - obter_namespace_mapi()
    - inicializar_mapi()
    - realizar_logon_mapi()
    - criar_conexao_outlook()
    - conectar_outlook()
    - verificar_outlook_disponivel()
    - verificar_outlook_aberto()
    - verificar_namespace_ativo()
    - obter_versao_outlook()
    - obter_nome_perfil()
    - obter_usuario_atual()
    - obter_email_usuario_atual()
    - obter_informacoes_conexao()
    - liberar_objeto_com()
    - desconectar_outlook()
Dependencias:
    - Outlook classico para Windows
    - pywin32
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de conexao a instancia ativa ou nova instancia.
        - Inclusao de inicializacao e logon MAPI.
        - Inclusao de consultas sobre perfil, usuario e versao.
        - Inclusao de objeto estruturado para armazenar a conexao.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa gc para solicitar a coleta de referencias COM liberadas
import gc

# Importa platform para validar o sistema operacional
import platform

# Importa dataclass para representar a conexao de forma estruturada
from dataclasses import dataclass

# Importa Any para tipar objetos COM dinamicos do Outlook
from typing import Any

# Importa a excecao base central quando disponivel
try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base utilizada quando o modulo central nao esta disponivel."""


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Identificador ProgID utilizado para automatizar o Outlook classico
OUTLOOK_PROG_ID = "Outlook.Application"

# O unico Namespace suportado pelo modelo de objetos do Outlook
NAMESPACE_MAPI = "MAPI"

# Pasta padrao Caixa de Entrada, utilizada para inicializar o MAPI
OL_FOLDER_INBOX = 6


# ----------------------------------------------------------------------------
# EXCECOES
# ----------------------------------------------------------------------------

class OutlookConexaoError(AutomacaoError):
    """Erro base para conexao e inicializacao do Outlook."""


class PyWin32NaoDisponivelError(OutlookConexaoError):
    """A biblioteca pywin32 nao esta instalada ou nao pode ser importada."""


class SistemaOperacionalIncompativelError(OutlookConexaoError):
    """A automacao COM foi solicitada fora do Windows."""


class OutlookNaoDisponivelError(OutlookConexaoError):
    """O Outlook classico nao esta instalado ou nao responde ao COM."""


class OutlookNaoAbertoError(OutlookConexaoError):
    """Nenhuma instancia ativa do Outlook foi encontrada."""


class NamespaceMAPIError(OutlookConexaoError):
    """O Namespace MAPI nao pode ser obtido ou inicializado."""


class PerfilOutlookError(OutlookConexaoError):
    """O perfil solicitado nao pode ser utilizado na sessao MAPI."""


class ObjetoCOMInvalidoError(OutlookConexaoError):
    """Um objeto COM obrigatorio esta ausente ou invalido."""


# ----------------------------------------------------------------------------
# MODELO DA CONEXAO
# ----------------------------------------------------------------------------

@dataclass(slots=True)
class ConexaoOutlook:
    """
    Agrupa os objetos COM principais de uma conexao com o Outlook.

    Attributes:
        aplicacao (Any):
            Objeto Outlook.Application.
        namespace (Any):
            Objeto Outlook.NameSpace do tipo MAPI.
        caixa_entrada (Any | None):
            Pasta Inbox usada para inicializar o MAPI.
        instancia_existente (bool):
            Indica se uma instancia aberta foi reutilizada.
        perfil (str | None):
            Perfil solicitado durante a conexao.
        inicializado (bool):
            Indica se o Namespace respondeu a uma pasta padrao.
    """

    aplicacao: Any
    namespace: Any
    caixa_entrada: Any | None = None
    instancia_existente: bool = False
    perfil: str | None = None
    inicializado: bool = False

    def para_dicionario(self) -> dict[str, Any]:
        """
        Retorna um resumo serializavel do estado da conexao.

        Returns:
            dict[str, Any]:
                Informacoes gerais sem expor diretamente objetos COM.
        """
        return {
            "instancia_existente": self.instancia_existente,
            "perfil": self.perfil,
            "inicializado": self.inicializado,
            "possui_aplicacao": self.aplicacao is not None,
            "possui_namespace": self.namespace is not None,
            "possui_caixa_entrada": self.caixa_entrada is not None,
        }

    def desconectar(self) -> None:
        """Libera as referencias COM armazenadas nesta conexao."""
        desconectar_outlook(self)


# ----------------------------------------------------------------------------
# VALIDACOES E IMPORTACAO
# ----------------------------------------------------------------------------

def _validar_objeto(objeto: Any, nome: str) -> None:
    """
    Valida se um objeto obrigatorio foi informado.

    Args:
        objeto (Any):
            Objeto que sera validado.
        nome (str):
            Nome utilizado na mensagem de erro.

    Returns:
        None:
            A funcao apenas realiza a validacao.
    """
    if objeto is None:
        raise ObjetoCOMInvalidoError(
            f"O parametro '{nome}' nao pode ser None."
        )


def _validar_booleano(valor: bool, nome: str) -> None:
    """Valida uma opcao booleana obrigatoria."""
    if not isinstance(valor, bool):
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")


def _validar_texto_opcional(valor: str | None, nome: str) -> str | None:
    """
    Valida e normaliza um texto opcional.

    Args:
        valor (str | None):
            Texto que sera validado.
        nome (str):
            Nome utilizado na mensagem de erro.

    Returns:
        str | None:
            Texto normalizado ou None.
    """
    if valor is None:
        return None

    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser string ou None.")

    texto = valor.strip()
    if not texto:
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")

    return texto


def verificar_pywin32_disponivel() -> bool:
    """
    Verifica se win32com.client pode ser importado.

    Returns:
        bool:
            True quando pywin32 estiver disponivel e False caso contrario.
    """
    try:
        import win32com.client  # type: ignore # noqa: F401
        return True
    except ImportError:
        return False


def importar_win32com() -> Any:
    """
    Importa e retorna o modulo win32com.client.

    Returns:
        Any:
            Modulo win32com.client importado.

    Raises:
        PyWin32NaoDisponivelError:
            Caso pywin32 nao esteja instalado.
    """
    try:
        import win32com.client as cliente  # type: ignore
        return cliente
    except ImportError as error:
        raise PyWin32NaoDisponivelError(
            "A biblioteca pywin32 nao esta instalada. "
            "Execute: python -m pip install pywin32"
        ) from error


def _validar_windows() -> None:
    """Valida se o codigo esta sendo executado no Windows."""
    if platform.system().lower() != "windows":
        raise SistemaOperacionalIncompativelError(
            "A automacao COM do Outlook classico requer Windows."
        )


# ----------------------------------------------------------------------------
# OBTENCAO DA APLICACAO
# ----------------------------------------------------------------------------

def obter_outlook_ativo(
    cliente_com: Any | None = None,
    validar_windows: bool = True,
) -> Any:
    """
    Retorna uma instancia do Outlook que ja esteja em execucao.

    Args:
        cliente_com (Any | None):
            Modulo win32com.client opcional, util para testes.
        validar_windows (bool):
            Define se o sistema operacional sera validado.

    Returns:
        Any:
            Objeto Outlook.Application ativo.

    Raises:
        OutlookNaoAbertoError:
            Caso nenhuma instancia ativa seja encontrada.
    """
    _validar_booleano(validar_windows, "validar_windows")
    if validar_windows:
        _validar_windows()

    cliente = cliente_com or importar_win32com()

    try:
        return cliente.GetActiveObject(OUTLOOK_PROG_ID)
    except Exception as error:
        raise OutlookNaoAbertoError(
            "Nenhuma instancia ativa do Outlook classico foi encontrada."
        ) from error


def criar_aplicacao_outlook(
    cliente_com: Any | None = None,
    usar_ensure_dispatch: bool = False,
    validar_windows: bool = True,
) -> Any:
    """
    Cria ou conecta a uma instancia COM do Outlook.

    Args:
        cliente_com (Any | None):
            Modulo win32com.client opcional.
        usar_ensure_dispatch (bool):
            Usa EnsureDispatch para gerar suporte de vinculacao antecipada.
        validar_windows (bool):
            Define se o sistema operacional sera validado.

    Returns:
        Any:
            Objeto Outlook.Application.
    """
    _validar_booleano(usar_ensure_dispatch, "usar_ensure_dispatch")
    _validar_booleano(validar_windows, "validar_windows")

    if validar_windows:
        _validar_windows()

    cliente = cliente_com or importar_win32com()

    try:
        if usar_ensure_dispatch:
            return cliente.gencache.EnsureDispatch(OUTLOOK_PROG_ID)
        return cliente.Dispatch(OUTLOOK_PROG_ID)
    except Exception as error:
        raise OutlookNaoDisponivelError(
            "Nao foi possivel criar a aplicacao do Outlook classico."
        ) from error


def obter_aplicacao_outlook(
    reutilizar_instancia: bool = True,
    criar_se_necessario: bool = True,
    cliente_com: Any | None = None,
    usar_ensure_dispatch: bool = False,
    validar_windows: bool = True,
) -> tuple[Any, bool]:
    """
    Obtem uma instancia ativa ou cria uma aplicacao do Outlook.

    Args:
        reutilizar_instancia (bool):
            Tenta reutilizar uma instancia aberta.
        criar_se_necessario (bool):
            Cria a aplicacao quando nenhuma instancia ativa existir.
        cliente_com (Any | None):
            Modulo win32com.client opcional.
        usar_ensure_dispatch (bool):
            Usa EnsureDispatch durante a criacao.
        validar_windows (bool):
            Define se o sistema operacional sera validado.

    Returns:
        tuple[Any, bool]:
            Aplicacao Outlook e indicador de instancia reutilizada.
    """
    for nome, valor in {
        "reutilizar_instancia": reutilizar_instancia,
        "criar_se_necessario": criar_se_necessario,
        "usar_ensure_dispatch": usar_ensure_dispatch,
        "validar_windows": validar_windows,
    }.items():
        _validar_booleano(valor, nome)

    # Tenta reutilizar uma instancia aberta
    if reutilizar_instancia:
        try:
            aplicacao = obter_outlook_ativo(
                cliente_com=cliente_com,
                validar_windows=validar_windows,
            )
            return aplicacao, True
        except OutlookNaoAbertoError:
            if not criar_se_necessario:
                raise

    # Impede criacao quando ela foi desativada
    if not criar_se_necessario:
        raise OutlookNaoAbertoError(
            "O Outlook nao esta aberto e a criacao de instancia foi desativada."
        )

    aplicacao = criar_aplicacao_outlook(
        cliente_com=cliente_com,
        usar_ensure_dispatch=usar_ensure_dispatch,
        validar_windows=validar_windows,
    )
    return aplicacao, False


# ----------------------------------------------------------------------------
# NAMESPACE E SESSAO MAPI
# ----------------------------------------------------------------------------

def obter_namespace_mapi(aplicacao: Any) -> Any:
    """
    Retorna o Namespace MAPI de uma aplicacao Outlook.

    Args:
        aplicacao (Any):
            Objeto Outlook.Application.

    Returns:
        Any:
            Objeto Outlook.NameSpace do tipo MAPI.
    """
    _validar_objeto(aplicacao, "aplicacao")

    try:
        # Algumas interfaces geradas expoem GetNamespace com s minusculo
        metodo = getattr(aplicacao, "GetNamespace", None)
        if callable(metodo):
            return metodo(NAMESPACE_MAPI)

        # Mantem compatibilidade com interfaces dinamicas que usam GetNameSpace
        metodo_alternativo = getattr(aplicacao, "GetNameSpace", None)
        if callable(metodo_alternativo):
            return metodo_alternativo(NAMESPACE_MAPI)

        # A propriedade Session e funcionalmente equivalente ao Namespace MAPI
        sessao = getattr(aplicacao, "Session", None)
        if sessao is not None:
            return sessao

        raise AttributeError("GetNamespace nao esta disponivel.")
    except Exception as error:
        raise NamespaceMAPIError(
            "Nao foi possivel obter o Namespace MAPI do Outlook."
        ) from error


def inicializar_mapi(
    namespace: Any,
    obter_caixa_entrada: bool = True,
) -> Any | None:
    """
    Inicializa o MAPI acessando uma pasta padrao do Outlook.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        obter_caixa_entrada (bool):
            Define se a pasta Inbox sera retornada.

    Returns:
        Any | None:
            Pasta Inbox ou None quando a consulta estiver desativada.
    """
    _validar_objeto(namespace, "namespace")
    _validar_booleano(obter_caixa_entrada, "obter_caixa_entrada")

    if not obter_caixa_entrada:
        return None

    try:
        return namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    except Exception as error:
        raise NamespaceMAPIError(
            "Nao foi possivel inicializar o MAPI pela Caixa de Entrada."
        ) from error


def realizar_logon_mapi(
    namespace: Any,
    perfil: str | None = None,
    exibir_dialogo: bool = False,
    nova_sessao: bool = False,
) -> Any:
    """
    Executa Logon no Namespace MAPI para um perfil especifico.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        perfil (str | None):
            Nome do perfil ou None para perfil padrao.
        exibir_dialogo (bool):
            Define se o seletor de perfil sera exibido.
        nova_sessao (bool):
            Solicita nova sessao MAPI.

    Returns:
        Any:
            O mesmo Namespace utilizado no logon.
    """
    _validar_objeto(namespace, "namespace")
    perfil_validado = _validar_texto_opcional(perfil, "perfil")
    _validar_booleano(exibir_dialogo, "exibir_dialogo")
    _validar_booleano(nova_sessao, "nova_sessao")

    try:
        namespace.Logon(
            perfil_validado or "",
            "",
            exibir_dialogo,
            nova_sessao,
        )
        return namespace
    except Exception as error:
        raise PerfilOutlookError(
            f"Nao foi possivel realizar o logon no perfil: "
            f"'{perfil_validado or 'padrao'}'."
        ) from error


def criar_conexao_outlook(
    aplicacao: Any,
    namespace: Any,
    caixa_entrada: Any | None = None,
    instancia_existente: bool = False,
    perfil: str | None = None,
) -> ConexaoOutlook:
    """
    Cria um objeto ConexaoOutlook a partir de objetos COM existentes.

    Returns:
        ConexaoOutlook:
            Conexao estruturada e validada.
    """
    _validar_objeto(aplicacao, "aplicacao")
    _validar_objeto(namespace, "namespace")
    _validar_booleano(instancia_existente, "instancia_existente")
    perfil_validado = _validar_texto_opcional(perfil, "perfil")

    return ConexaoOutlook(
        aplicacao=aplicacao,
        namespace=namespace,
        caixa_entrada=caixa_entrada,
        instancia_existente=instancia_existente,
        perfil=perfil_validado,
        inicializado=caixa_entrada is not None,
    )


def conectar_outlook(
    reutilizar_instancia: bool = True,
    criar_se_necessario: bool = True,
    inicializar: bool = True,
    perfil: str | None = None,
    realizar_logon: bool = False,
    exibir_dialogo: bool = False,
    nova_sessao: bool = False,
    cliente_com: Any | None = None,
    usar_ensure_dispatch: bool = False,
    validar_windows: bool = True,
) -> ConexaoOutlook:
    """
    Executa o fluxo completo de conexao com o Outlook classico.

    Returns:
        ConexaoOutlook:
            Aplicacao, Namespace e Inbox prontos para uso.
    """
    for nome, valor in {
        "inicializar": inicializar,
        "realizar_logon": realizar_logon,
        "exibir_dialogo": exibir_dialogo,
        "nova_sessao": nova_sessao,
    }.items():
        _validar_booleano(valor, nome)

    perfil_validado = _validar_texto_opcional(perfil, "perfil")

    # Obtem a aplicacao e informa se ela ja estava aberta
    aplicacao, instancia_existente = obter_aplicacao_outlook(
        reutilizar_instancia=reutilizar_instancia,
        criar_se_necessario=criar_se_necessario,
        cliente_com=cliente_com,
        usar_ensure_dispatch=usar_ensure_dispatch,
        validar_windows=validar_windows,
    )

    # Obtem o Namespace MAPI
    namespace = obter_namespace_mapi(aplicacao)

    # Logon explicito e indicado apenas quando solicitado
    if realizar_logon:
        realizar_logon_mapi(
            namespace,
            perfil=perfil_validado,
            exibir_dialogo=exibir_dialogo,
            nova_sessao=nova_sessao,
        )

    # Acessa a Inbox para inicializar o perfil padrao
    caixa_entrada = inicializar_mapi(namespace) if inicializar else None

    return criar_conexao_outlook(
        aplicacao=aplicacao,
        namespace=namespace,
        caixa_entrada=caixa_entrada,
        instancia_existente=instancia_existente,
        perfil=perfil_validado,
    )


# ----------------------------------------------------------------------------
# VERIFICACOES E INFORMACOES
# ----------------------------------------------------------------------------

def verificar_outlook_disponivel(
    cliente_com: Any | None = None,
    validar_windows: bool = True,
) -> bool:
    """Verifica se a aplicacao COM do Outlook pode ser criada."""
    try:
        criar_aplicacao_outlook(
            cliente_com=cliente_com,
            validar_windows=validar_windows,
        )
        return True
    except OutlookConexaoError:
        return False


def verificar_outlook_aberto(
    cliente_com: Any | None = None,
    validar_windows: bool = True,
) -> bool:
    """Verifica se existe uma instancia ativa do Outlook."""
    try:
        obter_outlook_ativo(
            cliente_com=cliente_com,
            validar_windows=validar_windows,
        )
        return True
    except OutlookConexaoError:
        return False


def verificar_namespace_ativo(namespace: Any) -> bool:
    """Verifica se o Namespace responde ao acesso de uma pasta padrao."""
    try:
        inicializar_mapi(namespace)
        return True
    except OutlookConexaoError:
        return False


def obter_versao_outlook(aplicacao: Any) -> str:
    """Retorna a versao informada pelo objeto Outlook.Application."""
    _validar_objeto(aplicacao, "aplicacao")
    try:
        return str(aplicacao.Version)
    except Exception as error:
        raise OutlookConexaoError(
            "Nao foi possivel obter a versao do Outlook."
        ) from error


def obter_nome_perfil(namespace: Any) -> str:
    """Retorna o nome do perfil ou sessao atual do Outlook."""
    _validar_objeto(namespace, "namespace")
    try:
        return str(namespace.CurrentProfileName)
    except Exception as error:
        raise PerfilOutlookError(
            "Nao foi possivel obter o nome do perfil atual."
        ) from error


def obter_usuario_atual(namespace: Any) -> Any:
    """Retorna o objeto Recipient representando o usuario atual."""
    _validar_objeto(namespace, "namespace")
    try:
        return namespace.CurrentUser
    except Exception as error:
        raise OutlookConexaoError(
            "Nao foi possivel obter o usuario atual do Outlook."
        ) from error


def obter_email_usuario_atual(namespace: Any) -> str:
    """
    Tenta obter o endereco SMTP do usuario atual do Outlook.

    Returns:
        str:
            Endereco SMTP ou endereco nativo disponibilizado pelo Outlook.
    """
    usuario = obter_usuario_atual(namespace)

    try:
        address_entry = usuario.AddressEntry

        # ExchangeUser disponibiliza PrimarySmtpAddress
        if address_entry is not None:
            tipo = str(getattr(address_entry, "Type", "")).upper()
            if tipo == "EX":
                exchange_user = address_entry.GetExchangeUser()
                smtp = getattr(exchange_user, "PrimarySmtpAddress", "")
                if smtp:
                    return str(smtp)

            endereco = getattr(address_entry, "Address", "")
            if endereco:
                return str(endereco)

        endereco_usuario = getattr(usuario, "Address", "")
        if endereco_usuario:
            return str(endereco_usuario)
    except Exception as error:
        raise OutlookConexaoError(
            "Nao foi possivel obter o e-mail do usuario atual."
        ) from error

    raise OutlookConexaoError(
        "O Outlook nao disponibilizou o e-mail do usuario atual."
    )


def obter_informacoes_conexao(
    conexao: ConexaoOutlook,
    incluir_email: bool = True,
) -> dict[str, Any]:
    """
    Retorna informacoes gerais da conexao ativa.

    Args:
        conexao (ConexaoOutlook):
            Conexao que sera consultada.
        incluir_email (bool):
            Define se o e-mail do usuario sera consultado.

    Returns:
        dict[str, Any]:
            Perfil, versao, usuario, e-mail e estado da conexao.
    """
    if not isinstance(conexao, ConexaoOutlook):
        raise TypeError(
            "O parametro 'conexao' deve ser uma instancia de ConexaoOutlook."
        )
    _validar_booleano(incluir_email, "incluir_email")

    usuario = obter_usuario_atual(conexao.namespace)
    dados = conexao.para_dicionario()
    dados.update(
        {
            "versao": obter_versao_outlook(conexao.aplicacao),
            "perfil_atual": obter_nome_perfil(conexao.namespace),
            "usuario": str(getattr(usuario, "Name", "")),
        }
    )

    if incluir_email:
        dados["email"] = obter_email_usuario_atual(conexao.namespace)

    return dados


# ----------------------------------------------------------------------------
# LIBERACAO DE REFERENCIAS
# ----------------------------------------------------------------------------

def liberar_objeto_com(objeto: Any) -> None:
    """
    Solicita a liberacao de uma referencia COM.

    Notes:
        A funcao nao encerra o Outlook. Ela apenas remove a referencia
        recebida e solicita uma coleta de lixo do Python.
    """
    # Objetos None nao exigem nenhuma operacao
    if objeto is None:
        return

    # Remove a referencia local e solicita a coleta de objetos sem uso
    del objeto
    gc.collect()


def desconectar_outlook(
    conexao: ConexaoOutlook,
    realizar_logoff: bool = False,
    encerrar_aplicacao: bool = False,
) -> None:
    """
    Libera objetos da conexao e opcionalmente executa Logoff ou Quit.

    Args:
        conexao (ConexaoOutlook):
            Conexao que sera encerrada.
        realizar_logoff (bool):
            Executa Namespace.Logoff.
        encerrar_aplicacao (bool):
            Executa Application.Quit. Use com cautela, pois fecha o Outlook.
    """
    if not isinstance(conexao, ConexaoOutlook):
        raise TypeError(
            "O parametro 'conexao' deve ser uma instancia de ConexaoOutlook."
        )

    _validar_booleano(realizar_logoff, "realizar_logoff")
    _validar_booleano(encerrar_aplicacao, "encerrar_aplicacao")

    # Mantem referencias temporarias para executar as acoes finais
    namespace = conexao.namespace
    aplicacao = conexao.aplicacao

    try:
        if realizar_logoff and namespace is not None:
            namespace.Logoff()

        if encerrar_aplicacao and aplicacao is not None:
            aplicacao.Quit()
    except Exception as error:
        raise OutlookConexaoError(
            "Nao foi possivel encerrar corretamente a conexao do Outlook."
        ) from error
    finally:
        # Remove referencias armazenadas na estrutura
        conexao.caixa_entrada = None
        conexao.namespace = None
        conexao.aplicacao = None
        conexao.inicializado = False

        liberar_objeto_com(namespace)
        liberar_objeto_com(aplicacao)
