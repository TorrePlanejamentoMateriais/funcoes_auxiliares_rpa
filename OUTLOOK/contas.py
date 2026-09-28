"""
============================================================
Modulo: OUTLOOK / contas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes auxiliares para consultar, localizar, validar e utilizar
    contas configuradas no perfil atual do Outlook classico via COM.
Funcoes disponiveis:
    - obter_colecao_contas()
    - contar_contas()
    - listar_contas()
    - obter_conta_por_indice()
    - obter_conta_por_email()
    - obter_conta_por_nome()
    - conta_existe()
    - obter_conta_padrao()
    - obter_email_conta()
    - obter_nome_conta()
    - obter_tipo_conta()
    - obter_usuario_conta()
    - obter_store_conta()
    - obter_pasta_raiz_conta()
    - obter_dados_conta()
    - definir_conta_envio()
    - identificar_conta_item()
    - listar_caixas_postais()
Dependencias:
    - Outlook classico para Windows
    - conexao.py ou um objeto Namespace MAPI valido
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de consulta e pesquisa de contas.
        - Inclusao de conta padrao e conta de envio.
        - Inclusao de extracao de dados e caixas postais.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa Any para tipar os objetos COM dinamicos do Outlook
from typing import Any

# Importa a excecao principal da biblioteca quando ela esta disponivel
try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base utilizada quando o modulo central nao esta disponivel."""

# Importa tipos de conta centralizados quando o modulo existe
try:
    from constantes import TipoConta
except ImportError:
    TipoConta = None


# ----------------------------------------------------------------------------
# EXCECOES
# ----------------------------------------------------------------------------

class OutlookContasError(AutomacaoError):
    """Erro base para operacoes relacionadas a contas do Outlook."""


class NamespaceInvalidoError(OutlookContasError):
    """O Namespace MAPI informado esta ausente ou nao possui Accounts."""


class ContaNaoEncontradaError(OutlookContasError):
    """Nenhuma conta correspondente foi encontrada no perfil atual."""


class ContaInvalidaError(OutlookContasError):
    """O objeto Account informado nao possui os dados esperados."""


class ContaEnvioError(OutlookContasError):
    """A conta de envio nao pode ser atribuida ao item do Outlook."""


class StoreContaError(OutlookContasError):
    """O Store ou a pasta raiz de uma conta nao pode ser obtido."""


# ----------------------------------------------------------------------------
# FUNCOES INTERNAS
# ----------------------------------------------------------------------------

def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """
    Obtem uma propriedade de um objeto COM sem interromper o fluxo.

    Args:
        objeto (Any):
            Objeto COM cuja propriedade sera consultada.
        nome (str):
            Nome da propriedade.
        padrao (Any):
            Valor retornado quando a propriedade nao estiver disponivel.

    Returns:
        Any:
            Valor da propriedade ou o valor padrao.
    """
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise ContaInvalidaError(
            f"O parametro '{nome}' nao pode ser None."
        )


def _validar_indice(indice: int, quantidade: int | None = None) -> None:
    """
    Valida um indice COM iniciado em um.

    Args:
        indice (int):
            Posicao da conta na colecao.
        quantidade (int | None):
            Quantidade maxima disponivel.
    """
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError("O parametro 'indice' deve ser inteiro.")

    if indice < 1:
        raise ValueError("O parametro 'indice' deve ser maior ou igual a 1.")

    if quantidade is not None and indice > quantidade:
        raise ContaNaoEncontradaError(
            f"Nao existe conta no indice {indice}. Quantidade: {quantidade}."
        )


def _validar_texto(valor: str, nome: str) -> str:
    """Valida e normaliza uma string obrigatoria."""
    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")

    texto = valor.strip()
    if not texto:
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")

    return texto


def _comparar_texto(
    atual: str,
    esperado: str,
    correspondencia_exata: bool,
    considerar_maiusculas: bool,
) -> bool:
    """Compara dois textos conforme as opcoes informadas."""
    if not isinstance(correspondencia_exata, bool):
        raise TypeError(
            "O parametro 'correspondencia_exata' deve ser booleano."
        )

    if not isinstance(considerar_maiusculas, bool):
        raise TypeError(
            "O parametro 'considerar_maiusculas' deve ser booleano."
        )

    valor_atual = atual if considerar_maiusculas else atual.lower()
    valor_esperado = esperado if considerar_maiusculas else esperado.lower()

    if correspondencia_exata:
        return valor_atual == valor_esperado

    return valor_esperado in valor_atual


def _normalizar_email(email: str) -> str:
    """Valida e normaliza um endereco utilizado na pesquisa de contas."""
    endereco = _validar_texto(email, "email").lower()

    if "@" not in endereco:
        raise ValueError(
            "O parametro 'email' deve possuir um endereco valido."
        )

    return endereco


# ----------------------------------------------------------------------------
# COLECAO DE CONTAS
# ----------------------------------------------------------------------------

def obter_colecao_contas(namespace: Any) -> Any:
    """
    Retorna a colecao Accounts do Namespace MAPI.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace ou Session.

    Returns:
        Any:
            Colecao COM Accounts.
    """
    if namespace is None:
        raise NamespaceInvalidoError(
            "O Namespace MAPI nao pode ser None."
        )

    try:
        contas = namespace.Accounts
    except Exception as error:
        raise NamespaceInvalidoError(
            "O Namespace informado nao possui uma colecao Accounts valida."
        ) from error

    if contas is None:
        raise NamespaceInvalidoError(
            "A colecao Accounts nao esta disponivel."
        )

    return contas


def contar_contas(namespace: Any) -> int:
    """
    Retorna a quantidade de contas configuradas no perfil atual.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.

    Returns:
        int:
            Quantidade de contas da colecao.
    """
    contas = obter_colecao_contas(namespace)

    try:
        return int(contas.Count)
    except Exception as error:
        raise NamespaceInvalidoError(
            "Nao foi possivel contar as contas do perfil."
        ) from error


def obter_conta_por_indice(namespace: Any, indice: int) -> Any:
    """
    Retorna a conta localizada em determinado indice COM.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        indice (int):
            Posicao iniciada em um.

    Returns:
        Any:
            Objeto Account localizado.
    """
    quantidade = contar_contas(namespace)
    _validar_indice(indice, quantidade)

    try:
        return obter_colecao_contas(namespace).Item(indice)
    except Exception as error:
        raise ContaNaoEncontradaError(
            f"Nao foi possivel obter a conta no indice {indice}."
        ) from error


def listar_contas(namespace: Any) -> list[dict[str, Any]]:
    """
    Retorna os dados das contas configuradas no perfil atual.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.

    Returns:
        list[dict[str, Any]]:
            Dados estruturados das contas na ordem da colecao.
    """
    resultado = []

    for indice in range(1, contar_contas(namespace) + 1):
        conta = obter_conta_por_indice(namespace, indice)
        dados = obter_dados_conta(conta)
        dados["indice"] = indice
        resultado.append(dados)

    return resultado


# ----------------------------------------------------------------------------
# CONSULTAS DE PROPRIEDADES
# ----------------------------------------------------------------------------

def obter_email_conta(conta: Any, permitir_vazio: bool = False) -> str:
    """
    Retorna o endereco SMTP configurado na conta.

    Args:
        conta (Any):
            Objeto Outlook.Account.
        permitir_vazio (bool):
            Permite retornar string vazia quando nao houver SMTP.

    Returns:
        str:
            Endereco SMTP da conta.
    """
    _validar_objeto(conta, "conta")

    if not isinstance(permitir_vazio, bool):
        raise TypeError("O parametro 'permitir_vazio' deve ser booleano.")

    smtp = str(_obter_propriedade(conta, "SmtpAddress", "") or "").strip()

    if not smtp and not permitir_vazio:
        raise ContaInvalidaError(
            "A conta nao possui um endereco SMTP disponivel."
        )

    return smtp


def obter_nome_conta(conta: Any) -> str:
    """
    Retorna o nome amigavel de exibicao da conta.

    Args:
        conta (Any):
            Objeto Outlook.Account.

    Returns:
        str:
            DisplayName, UserName ou SMTP da conta.
    """
    _validar_objeto(conta, "conta")

    for propriedade in ("DisplayName", "UserName", "SmtpAddress"):
        valor = str(_obter_propriedade(conta, propriedade, "") or "").strip()
        if valor:
            return valor

    raise ContaInvalidaError(
        "A conta nao possui nome, usuario ou endereco SMTP."
    )


def obter_tipo_conta(conta: Any, converter_enum: bool = False) -> Any:
    """
    Retorna o tipo numerico da conta ou a enumeracao TipoConta.

    Args:
        conta (Any):
            Objeto Outlook.Account.
        converter_enum (bool):
            Converte o valor para TipoConta quando possivel.

    Returns:
        Any:
            Inteiro ou membro de TipoConta.
    """
    _validar_objeto(conta, "conta")

    if not isinstance(converter_enum, bool):
        raise TypeError("O parametro 'converter_enum' deve ser booleano.")

    try:
        tipo = int(conta.AccountType)
    except Exception as error:
        raise ContaInvalidaError(
            "Nao foi possivel obter o tipo da conta."
        ) from error

    if converter_enum and TipoConta is not None:
        try:
            return TipoConta(tipo)
        except ValueError:
            return tipo

    return tipo


def obter_usuario_conta(conta: Any) -> str:
    """Retorna a propriedade UserName de uma conta Outlook."""
    _validar_objeto(conta, "conta")
    usuario = str(_obter_propriedade(conta, "UserName", "") or "").strip()

    if not usuario:
        raise ContaInvalidaError(
            "A conta nao possui um nome de usuario disponivel."
        )

    return usuario


def obter_store_conta(conta: Any) -> Any:
    """
    Retorna o Store associado a uma conta Outlook.

    Args:
        conta (Any):
            Objeto Outlook.Account.

    Returns:
        Any:
            Objeto Store da conta.
    """
    _validar_objeto(conta, "conta")

    try:
        store = conta.DeliveryStore
    except Exception as error:
        raise StoreContaError(
            "Nao foi possivel obter o DeliveryStore da conta."
        ) from error

    if store is None:
        raise StoreContaError(
            "A conta nao possui um DeliveryStore disponivel."
        )

    return store


def obter_pasta_raiz_conta(conta: Any) -> Any:
    """
    Retorna a pasta raiz da caixa postal vinculada a conta.

    Args:
        conta (Any):
            Objeto Outlook.Account.

    Returns:
        Any:
            Objeto Folder raiz do Store.
    """
    store = obter_store_conta(conta)

    try:
        pasta = store.GetRootFolder()
    except Exception as error:
        raise StoreContaError(
            "Nao foi possivel obter a pasta raiz da conta."
        ) from error

    if pasta is None:
        raise StoreContaError(
            "A pasta raiz da conta nao esta disponivel."
        )

    return pasta


def obter_dados_conta(conta: Any) -> dict[str, Any]:
    """
    Converte as principais propriedades de uma conta para dicionario.

    Args:
        conta (Any):
            Objeto Outlook.Account.

    Returns:
        dict[str, Any]:
            Nome, usuario, SMTP, tipo e informacoes do Store.
    """
    _validar_objeto(conta, "conta")

    store = _obter_propriedade(conta, "DeliveryStore")

    return {
        "nome": str(_obter_propriedade(conta, "DisplayName", "") or ""),
        "usuario": str(_obter_propriedade(conta, "UserName", "") or ""),
        "email": str(_obter_propriedade(conta, "SmtpAddress", "") or ""),
        "tipo": _obter_propriedade(conta, "AccountType"),
        "store_id": _obter_propriedade(store, "StoreID"),
        "store_nome": str(_obter_propriedade(store, "DisplayName", "") or ""),
        "exchange_connection_mode": _obter_propriedade(
            conta,
            "ExchangeConnectionMode",
        ),
        "exchange_mailbox_server_name": str(
            _obter_propriedade(
                conta,
                "ExchangeMailboxServerName",
                "",
            ) or ""
        ),
    }


# ----------------------------------------------------------------------------
# PESQUISA DE CONTAS
# ----------------------------------------------------------------------------

def obter_conta_por_email(
    namespace: Any,
    email: str,
    considerar_maiusculas: bool = False,
) -> Any:
    """
    Localiza uma conta pelo endereco SMTP.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        email (str):
            Endereco SMTP procurado.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia letras maiusculas.

    Returns:
        Any:
            Objeto Account correspondente.
    """
    endereco = _normalizar_email(email)

    if not isinstance(considerar_maiusculas, bool):
        raise TypeError(
            "O parametro 'considerar_maiusculas' deve ser booleano."
        )

    if considerar_maiusculas:
        endereco = email.strip()

    for indice in range(1, contar_contas(namespace) + 1):
        conta = obter_conta_por_indice(namespace, indice)
        smtp = obter_email_conta(conta, permitir_vazio=True)

        if considerar_maiusculas:
            corresponde = smtp == endereco
        else:
            corresponde = smtp.lower() == endereco.lower()

        if corresponde:
            return conta

    raise ContaNaoEncontradaError(
        f"Nenhuma conta foi encontrada para o e-mail '{email}'."
    )


def obter_conta_por_nome(
    namespace: Any,
    nome: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
) -> Any:
    """
    Localiza uma conta pelo nome de exibicao ou nome de usuario.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        nome (str):
            Nome completo ou parcial procurado.
        correspondencia_exata (bool):
            Define se a comparacao deve ser exata.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia letras maiusculas.

    Returns:
        Any:
            Primeiro objeto Account correspondente.
    """
    texto = _validar_texto(nome, "nome")

    for indice in range(1, contar_contas(namespace) + 1):
        conta = obter_conta_por_indice(namespace, indice)
        candidatos = [
            str(_obter_propriedade(conta, "DisplayName", "") or ""),
            str(_obter_propriedade(conta, "UserName", "") or ""),
            str(_obter_propriedade(conta, "SmtpAddress", "") or ""),
        ]

        if any(
            _comparar_texto(
                candidato,
                texto,
                correspondencia_exata,
                considerar_maiusculas,
            )
            for candidato in candidatos
            if candidato
        ):
            return conta

    raise ContaNaoEncontradaError(
        f"Nenhuma conta correspondente a '{nome}' foi encontrada."
    )


def conta_existe(
    namespace: Any,
    identificador: str,
    buscar_por_email: bool = True,
) -> bool:
    """
    Verifica se uma conta existe por e-mail ou nome.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        identificador (str):
            E-mail ou nome procurado.
        buscar_por_email (bool):
            Usa pesquisa SMTP quando True e pesquisa por nome quando False.

    Returns:
        bool:
            True quando a conta existir e False caso contrario.
    """
    if not isinstance(buscar_por_email, bool):
        raise TypeError("O parametro 'buscar_por_email' deve ser booleano.")

    try:
        if buscar_por_email:
            obter_conta_por_email(namespace, identificador)
        else:
            obter_conta_por_nome(namespace, identificador)
        return True
    except ContaNaoEncontradaError:
        return False


def obter_conta_padrao(namespace: Any) -> Any:
    """
    Retorna a conta padrao utilizada pelo perfil atual.

    Notes:
        A colecao Accounts nao expoe uma propriedade universal de conta
        padrao. A funcao tenta Session.CurrentUser e, como fallback,
        retorna a primeira conta da colecao.
    """
    _validar_objeto(namespace, "namespace")

    # Tenta correlacionar o usuario atual com uma conta SMTP
    try:
        usuario = namespace.CurrentUser
        address_entry = _obter_propriedade(usuario, "AddressEntry")
        email = ""

        if address_entry is not None:
            tipo = str(_obter_propriedade(address_entry, "Type", "")).upper()
            if tipo == "EX":
                exchange_user = address_entry.GetExchangeUser()
                email = str(
                    _obter_propriedade(
                        exchange_user,
                        "PrimarySmtpAddress",
                        "",
                    ) or ""
                )
            if not email:
                email = str(
                    _obter_propriedade(address_entry, "Address", "") or ""
                )

        if not email:
            email = str(_obter_propriedade(usuario, "Address", "") or "")

        if email and "@" in email:
            return obter_conta_por_email(namespace, email)
    except (ContaNaoEncontradaError, Exception):
        pass

    # Usa a primeira conta como fallback consistente
    if contar_contas(namespace) < 1:
        raise ContaNaoEncontradaError(
            "O perfil atual nao possui contas configuradas."
        )

    return obter_conta_por_indice(namespace, 1)


# ----------------------------------------------------------------------------
# USO DA CONTA EM ITENS E STORES
# ----------------------------------------------------------------------------

def definir_conta_envio(item: Any, conta: Any) -> Any:
    """
    Define a propriedade SendUsingAccount de um item do Outlook.

    Args:
        item (Any):
            MailItem, AppointmentItem ou outro item enviavel.
        conta (Any):
            Objeto Outlook.Account utilizado no envio.

    Returns:
        Any:
            O mesmo item configurado.
    """
    _validar_objeto(item, "item")
    _validar_objeto(conta, "conta")

    try:
        item.SendUsingAccount = conta
        return item
    except Exception as error:
        raise ContaEnvioError(
            "Nao foi possivel definir a conta de envio do item."
        ) from error


def identificar_conta_item(
    namespace: Any,
    item: Any,
) -> Any:
    """
    Tenta identificar a conta associada a um item do Outlook.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.
        item (Any):
            Item que sera analisado.

    Returns:
        Any:
            Conta encontrada por SendUsingAccount ou StoreID.
    """
    _validar_objeto(item, "item")

    # Itens em composicao podem possuir SendUsingAccount
    conta_envio = _obter_propriedade(item, "SendUsingAccount")
    if conta_envio is not None:
        return conta_envio

    # Itens salvos normalmente possuem Parent.StoreID
    parent = _obter_propriedade(item, "Parent")
    store_id_item = _obter_propriedade(parent, "StoreID")

    if store_id_item:
        for indice in range(1, contar_contas(namespace) + 1):
            conta = obter_conta_por_indice(namespace, indice)
            store = _obter_propriedade(conta, "DeliveryStore")
            store_id_conta = _obter_propriedade(store, "StoreID")

            if store_id_conta == store_id_item:
                return conta

    raise ContaNaoEncontradaError(
        "Nao foi possivel identificar a conta associada ao item."
    )


def listar_caixas_postais(namespace: Any) -> list[dict[str, Any]]:
    """
    Retorna as pastas raiz das contas configuradas no perfil.

    Args:
        namespace (Any):
            Objeto Outlook.NameSpace.

    Returns:
        list[dict[str, Any]]:
            Conta, e-mail, StoreID, pasta raiz e caminho da pasta.
    """
    resultado = []

    for indice in range(1, contar_contas(namespace) + 1):
        conta = obter_conta_por_indice(namespace, indice)

        try:
            store = obter_store_conta(conta)
            pasta_raiz = obter_pasta_raiz_conta(conta)
        except StoreContaError:
            # Contas sem Store local nao representam caixas postais navegaveis
            continue

        resultado.append(
            {
                "indice": indice,
                "conta": conta,
                "nome": obter_nome_conta(conta),
                "email": obter_email_conta(conta, permitir_vazio=True),
                "store": store,
                "store_id": _obter_propriedade(store, "StoreID"),
                "pasta_raiz": pasta_raiz,
                "nome_pasta_raiz": str(
                    _obter_propriedade(pasta_raiz, "Name", "") or ""
                ),
                "caminho_pasta": str(
                    _obter_propriedade(pasta_raiz, "FolderPath", "") or ""
                ),
            }
        )

    return resultado
