"""
============================================================
Modulo: OUTLOOK / emails.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes reutilizaveis para criar, configurar, salvar, exibir,
    enviar, copiar, mover e extrair dados de e-mails do Outlook classico.
Dependencias:
    - Outlook classico para Windows
    - constantes.py
    - anexos.py (opcional, para adicionar arquivos)
    - contas.py (opcional, para selecionar conta de envio)
============================================================
"""

from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base usada quando o modulo central nao esta disponivel."""

try:
    from constantes import (
        OL_BCC,
        OL_CC,
        OL_FORMAT_HTML,
        OL_FORMAT_PLAIN,
        OL_IMPORTANCE_HIGH,
        OL_IMPORTANCE_NORMAL,
        OL_MAIL_ITEM,
        OL_TO,
    )
except ImportError:
    OL_MAIL_ITEM = 0
    OL_FORMAT_PLAIN = 1
    OL_FORMAT_HTML = 2
    OL_IMPORTANCE_NORMAL = 1
    OL_IMPORTANCE_HIGH = 2
    OL_TO = 1
    OL_CC = 2
    OL_BCC = 3


class OutlookEmailsError(AutomacaoError):
    """Erro base para operacoes relacionadas a e-mails do Outlook."""


class EmailInvalidoError(OutlookEmailsError):
    """O MailItem informado esta ausente ou possui dados invalidos."""


class EmailCriacaoError(OutlookEmailsError):
    """O Outlook nao conseguiu criar um novo MailItem."""


class EmailEnvioError(OutlookEmailsError):
    """O Outlook nao conseguiu enviar o e-mail."""


class EmailSalvamentoError(OutlookEmailsError):
    """O Outlook nao conseguiu salvar o e-mail."""


class DestinatarioInvalidoError(OutlookEmailsError):
    """Um ou mais destinatarios sao invalidos ou nao foram resolvidos."""


class EmailAcaoError(OutlookEmailsError):
    """Uma acao como mover, copiar, excluir ou exibir falhou."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM, retornando o padrao em caso de erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise EmailInvalidoError(f"O parametro '{nome}' nao pode ser None.")


def _validar_texto(valor: str, nome: str, permitir_vazio: bool = False) -> str:
    """Valida e normaliza uma propriedade textual."""
    if not isinstance(permitir_vazio, bool):
        raise TypeError("O parametro 'permitir_vazio' deve ser booleano.")
    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")
    texto = valor.strip()
    if not texto and not permitir_vazio:
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")
    return texto


def _normalizar_destinatarios(
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
) -> list[str]:
    """Normaliza um ou varios destinatarios, removendo duplicidades."""
    if isinstance(destinatarios, str):
        destinatarios = [p for p in destinatarios.replace(",", ";").split(";") if p.strip()]
    if not isinstance(destinatarios, (list, tuple, set)) or not destinatarios:
        raise TypeError("Destinatarios deve ser string ou colecao nao vazia.")
    resultado = []
    for destinatario in destinatarios:
        texto = _validar_texto(destinatario, "destinatario")
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def _validar_booleano(valor: bool, nome: str) -> None:
    """Valida uma opcao booleana."""
    if not isinstance(valor, bool):
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")


def _definir_propriedade(email: Any, propriedade: str, valor: Any) -> None:
    """Define uma propriedade de MailItem com tratamento padronizado."""
    _validar_objeto(email, "email")
    try:
        setattr(email, propriedade, valor)
    except Exception as error:
        raise EmailInvalidoError(
            f"Nao foi possivel definir a propriedade '{propriedade}'."
        ) from error


def criar_email(outlook: Any) -> Any:
    """Cria e retorna um novo MailItem ainda nao salvo."""
    _validar_objeto(outlook, "outlook")
    try:
        return outlook.CreateItem(OL_MAIL_ITEM)
    except Exception as error:
        raise EmailCriacaoError("Nao foi possivel criar o e-mail.") from error


def definir_assunto(email: Any, assunto: str) -> Any:
    """Define o assunto do e-mail."""
    _definir_propriedade(email, "Subject", _validar_texto(assunto, "assunto"))
    return email


def definir_corpo_texto(email: Any, corpo: str) -> Any:
    """Define o corpo em texto simples."""
    _definir_propriedade(email, "BodyFormat", OL_FORMAT_PLAIN)
    _definir_propriedade(email, "Body", _validar_texto(corpo, "corpo", True))
    return email


def definir_corpo_html(email: Any, html: str) -> Any:
    """Define o corpo HTML do e-mail."""
    _definir_propriedade(email, "BodyFormat", OL_FORMAT_HTML)
    _definir_propriedade(email, "HTMLBody", _validar_texto(html, "html", True))
    return email


def adicionar_assinatura_html(email: Any, assinatura_html: str, antes: bool = False) -> Any:
    """Adiciona HTML ao corpo atual, antes ou depois do conteudo existente."""
    _validar_booleano(antes, "antes")
    assinatura = _validar_texto(assinatura_html, "assinatura_html", True)
    atual = str(_obter_propriedade(email, "HTMLBody", "") or "")
    definir_corpo_html(email, assinatura + atual if antes else atual + assinatura)
    return email


def definir_destinatarios(
    email: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define o campo Para utilizando uma lista separada por ponto e virgula."""
    _definir_propriedade(email, "To", "; ".join(_normalizar_destinatarios(destinatarios)))
    return email


def definir_copias(
    email: Any,
    copias: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define o campo CC."""
    _definir_propriedade(email, "CC", "; ".join(_normalizar_destinatarios(copias)))
    return email


def definir_copias_ocultas(
    email: Any,
    copias_ocultas: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define o campo CCO."""
    _definir_propriedade(email, "BCC", "; ".join(_normalizar_destinatarios(copias_ocultas)))
    return email


def adicionar_destinatario(email: Any, endereco: str, tipo: int = OL_TO) -> Any:
    """Adiciona um Recipient do tipo Para, CC ou CCO."""
    _validar_objeto(email, "email")
    endereco = _validar_texto(endereco, "endereco")
    if isinstance(tipo, bool) or not isinstance(tipo, int):
        raise TypeError("O parametro 'tipo' deve ser inteiro.")
    if tipo not in {OL_TO, OL_CC, OL_BCC}:
        raise ValueError("O tipo deve representar Para, CC ou CCO.")
    try:
        destinatario = email.Recipients.Add(endereco)
        destinatario.Type = tipo
        return destinatario
    except Exception as error:
        raise DestinatarioInvalidoError(
            f"Nao foi possivel adicionar o destinatario '{endereco}'."
        ) from error


def adicionar_destinatarios(
    email: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    tipo: int = OL_TO,
) -> list[Any]:
    """Adiciona varios destinatarios pela colecao Recipients."""
    return [
        adicionar_destinatario(email, endereco, tipo)
        for endereco in _normalizar_destinatarios(destinatarios)
    ]


def resolver_destinatarios(email: Any, exigir_resolucao: bool = True) -> bool:
    """Solicita ao Outlook a resolucao de todos os destinatarios."""
    _validar_objeto(email, "email")
    _validar_booleano(exigir_resolucao, "exigir_resolucao")
    try:
        resolvido = bool(email.Recipients.ResolveAll())
    except Exception as error:
        raise DestinatarioInvalidoError(
            "Nao foi possivel resolver os destinatarios."
        ) from error
    if exigir_resolucao and not resolvido:
        raise DestinatarioInvalidoError(
            "Um ou mais destinatarios nao foram resolvidos."
        )
    return resolvido


def listar_destinatarios(email: Any) -> list[dict[str, Any]]:
    """Retorna os destinatarios e seus metadados principais."""
    _validar_objeto(email, "email")
    try:
        colecao = email.Recipients
        quantidade = int(colecao.Count)
    except Exception as error:
        raise DestinatarioInvalidoError(
            "Nao foi possivel consultar os destinatarios."
        ) from error
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            destinatario = colecao.Item(indice)
            entrada = _obter_propriedade(destinatario, "AddressEntry")
            resultado.append({
                "indice": indice,
                "nome": str(_obter_propriedade(destinatario, "Name", "") or ""),
                "endereco": str(_obter_propriedade(entrada, "Address", "") or ""),
                "tipo": _obter_propriedade(destinatario, "Type"),
                "resolvido": bool(_obter_propriedade(destinatario, "Resolved", False)),
            })
        except Exception as error:
            raise DestinatarioInvalidoError(
                f"Nao foi possivel consultar o destinatario {indice}."
            ) from error
    return resultado


def definir_importancia(email: Any, importancia: int = OL_IMPORTANCE_NORMAL) -> Any:
    """Define a importancia baixa, normal ou alta do e-mail."""
    if isinstance(importancia, bool) or not isinstance(importancia, int):
        raise TypeError("O parametro 'importancia' deve ser inteiro.")
    if importancia not in {0, 1, 2}:
        raise ValueError("A importancia deve ser 0, 1 ou 2.")
    _definir_propriedade(email, "Importance", importancia)
    return email


def definir_confirmacoes(
    email: Any,
    confirmacao_leitura: bool = False,
    confirmacao_entrega: bool = False,
) -> Any:
    """Configura confirmacoes de leitura e entrega."""
    _validar_booleano(confirmacao_leitura, "confirmacao_leitura")
    _validar_booleano(confirmacao_entrega, "confirmacao_entrega")
    _definir_propriedade(email, "ReadReceiptRequested", confirmacao_leitura)
    _definir_propriedade(email, "OriginatorDeliveryReportRequested", confirmacao_entrega)
    return email


def definir_envio_agendado(email: Any, data_envio: datetime | None) -> Any:
    """Define DeferredDeliveryTime ou remove o agendamento com None."""
    if data_envio is not None and not isinstance(data_envio, datetime):
        raise TypeError("O parametro 'data_envio' deve ser datetime ou None.")
    _definir_propriedade(email, "DeferredDeliveryTime", data_envio)
    return email


def definir_categorias(
    email: Any,
    categorias: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define categorias sem duplicidades."""
    valores = _normalizar_destinatarios(categorias)
    _definir_propriedade(email, "Categories", ", ".join(valores))
    return email


def definir_conta_envio(email: Any, conta: Any) -> Any:
    """Define SendUsingAccount para selecionar a conta de envio."""
    _validar_objeto(email, "email")
    _validar_objeto(conta, "conta")
    _definir_propriedade(email, "SendUsingAccount", conta)
    return email


def adicionar_anexo(email: Any, caminho: str | Path, nome_exibicao: str | None = None) -> Any:
    """Adiciona um arquivo existente a colecao Attachments."""
    _validar_objeto(email, "email")
    if not isinstance(caminho, (str, Path)):
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    arquivo = Path(caminho).expanduser().resolve()
    if not arquivo.is_file():
        raise FileNotFoundError(f"O arquivo nao existe: {arquivo}")
    nome = arquivo.name if nome_exibicao is None else _validar_texto(nome_exibicao, "nome_exibicao")
    try:
        return email.Attachments.Add(str(arquivo), 1, 1, nome)
    except Exception as error:
        raise EmailInvalidoError(
            f"Nao foi possivel adicionar o anexo: {arquivo}"
        ) from error


def adicionar_anexos(
    email: Any,
    caminhos: list[str | Path] | tuple[str | Path, ...] | set[str | Path],
    ignorar_inexistentes: bool = False,
) -> list[Any]:
    """Adiciona varios anexos ao e-mail."""
    if not isinstance(caminhos, (list, tuple, set)) or not caminhos:
        raise TypeError("O parametro 'caminhos' deve ser uma colecao nao vazia.")
    _validar_booleano(ignorar_inexistentes, "ignorar_inexistentes")
    resultado = []
    for caminho in caminhos:
        arquivo = Path(caminho).expanduser().resolve()
        if ignorar_inexistentes and not arquivo.is_file():
            continue
        resultado.append(adicionar_anexo(email, arquivo))
    return resultado


def configurar_email(
    email: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    assunto: str,
    corpo: str,
    html: bool = False,
    copias: str | list[str] | tuple[str, ...] | set[str] | None = None,
    copias_ocultas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    anexos: list[str | Path] | tuple[str | Path, ...] | set[str | Path] | None = None,
    importancia: int = OL_IMPORTANCE_NORMAL,
    conta_envio: Any | None = None,
    data_envio: datetime | None = None,
) -> Any:
    """Configura as principais propriedades de um e-mail."""
    _validar_booleano(html, "html")
    definir_destinatarios(email, destinatarios)
    definir_assunto(email, assunto)
    definir_corpo_html(email, corpo) if html else definir_corpo_texto(email, corpo)
    definir_importancia(email, importancia)
    if copias is not None:
        definir_copias(email, copias)
    if copias_ocultas is not None:
        definir_copias_ocultas(email, copias_ocultas)
    if anexos is not None:
        adicionar_anexos(email, anexos)
    if conta_envio is not None:
        definir_conta_envio(email, conta_envio)
    if data_envio is not None:
        definir_envio_agendado(email, data_envio)
    return email


def criar_email_completo(
    outlook: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    assunto: str,
    corpo: str,
    enviar: bool = False,
    exibir: bool = False,
    salvar: bool = False,
    **configuracoes: Any,
) -> Any:
    """Cria, configura e opcionalmente salva, exibe ou envia um e-mail."""
    for nome, valor in {"enviar": enviar, "exibir": exibir, "salvar": salvar}.items():
        _validar_booleano(valor, nome)
    email = criar_email(outlook)
    configurar_email(email, destinatarios, assunto, corpo, **configuracoes)
    if salvar:
        salvar_rascunho(email)
    if exibir:
        exibir_email(email)
    if enviar:
        enviar_email(email)
    return email


def salvar_rascunho(email: Any) -> Any:
    """Salva o e-mail e retorna o proprio MailItem."""
    _validar_objeto(email, "email")
    try:
        email.Save()
        return email
    except Exception as error:
        raise EmailSalvamentoError("Nao foi possivel salvar o e-mail.") from error


def exibir_email(email: Any, modal: bool = False) -> Any:
    """Exibe a janela do e-mail no Outlook."""
    _validar_objeto(email, "email")
    _validar_booleano(modal, "modal")
    try:
        email.Display(modal)
        return email
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel exibir o e-mail.") from error


def enviar_email(email: Any, resolver: bool = True) -> Any:
    """Resolve destinatarios e envia o MailItem."""
    _validar_objeto(email, "email")
    _validar_booleano(resolver, "resolver")
    if resolver:
        resolver_destinatarios(email, exigir_resolucao=True)
    try:
        email.Send()
        return email
    except Exception as error:
        raise EmailEnvioError("Nao foi possivel enviar o e-mail.") from error


def enviar_email_simples(
    outlook: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    assunto: str,
    corpo: str,
    **configuracoes: Any,
) -> Any:
    """Cria e envia imediatamente um e-mail em texto simples."""
    return criar_email_completo(
        outlook, destinatarios, assunto, corpo,
        enviar=True, html=False, **configuracoes,
    )


def enviar_email_html(
    outlook: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    assunto: str,
    html: str,
    **configuracoes: Any,
) -> Any:
    """Cria e envia imediatamente um e-mail HTML."""
    return criar_email_completo(
        outlook, destinatarios, assunto, html,
        enviar=True, html=True, **configuracoes,
    )


def responder_email(email_original: Any, responder_todos: bool = False) -> Any:
    """Cria uma resposta simples ou uma resposta a todos."""
    _validar_objeto(email_original, "email_original")
    _validar_booleano(responder_todos, "responder_todos")
    try:
        return email_original.ReplyAll() if responder_todos else email_original.Reply()
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel criar a resposta.") from error


def encaminhar_email(email_original: Any) -> Any:
    """Cria um novo e-mail encaminhando a mensagem original."""
    _validar_objeto(email_original, "email_original")
    try:
        return email_original.Forward()
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel encaminhar o e-mail.") from error


def copiar_email(email: Any) -> Any:
    """Cria e retorna uma copia do MailItem."""
    _validar_objeto(email, "email")
    try:
        return email.Copy()
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel copiar o e-mail.") from error


def mover_email(email: Any, pasta_destino: Any) -> Any:
    """Move o e-mail para outra pasta e retorna o item movido."""
    _validar_objeto(email, "email")
    _validar_objeto(pasta_destino, "pasta_destino")
    try:
        return email.Move(pasta_destino)
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel mover o e-mail.") from error


def excluir_email(email: Any) -> bool:
    """Exclui o e-mail da pasta atual."""
    _validar_objeto(email, "email")
    try:
        email.Delete()
        return True
    except Exception as error:
        raise EmailAcaoError("Nao foi possivel excluir o e-mail.") from error


def marcar_como_lido(email: Any, lido: bool = True, salvar: bool = False) -> Any:
    """Altera o estado UnRead e opcionalmente salva o item."""
    _validar_booleano(lido, "lido")
    _validar_booleano(salvar, "salvar")
    _definir_propriedade(email, "UnRead", not lido)
    if salvar:
        salvar_rascunho(email)
    return email


def obter_dados_email(email: Any) -> dict[str, Any]:
    """Converte as principais propriedades de MailItem para dicionario."""
    _validar_objeto(email, "email")
    anexos = _obter_propriedade(email, "Attachments")
    return {
        "entry_id": _obter_propriedade(email, "EntryID"),
        "assunto": str(_obter_propriedade(email, "Subject", "") or ""),
        "remetente_nome": str(_obter_propriedade(email, "SenderName", "") or ""),
        "remetente_email": str(_obter_propriedade(email, "SenderEmailAddress", "") or ""),
        "para": str(_obter_propriedade(email, "To", "") or ""),
        "cc": str(_obter_propriedade(email, "CC", "") or ""),
        "cco": str(_obter_propriedade(email, "BCC", "") or ""),
        "corpo": str(_obter_propriedade(email, "Body", "") or ""),
        "corpo_html": str(_obter_propriedade(email, "HTMLBody", "") or ""),
        "recebido_em": _obter_propriedade(email, "ReceivedTime"),
        "enviado_em": _obter_propriedade(email, "SentOn"),
        "criado_em": _obter_propriedade(email, "CreationTime"),
        "nao_lido": bool(_obter_propriedade(email, "UnRead", False)),
        "importancia": _obter_propriedade(email, "Importance"),
        "categorias": str(_obter_propriedade(email, "Categories", "") or ""),
        "conversation_id": _obter_propriedade(email, "ConversationID"),
        "message_class": str(_obter_propriedade(email, "MessageClass", "") or ""),
        "quantidade_anexos": int(_obter_propriedade(anexos, "Count", 0) or 0),
    }
