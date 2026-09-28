"""
============================================================
Modulo: OUTLOOK / respostas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes reutilizaveis para responder, responder a todos e encaminhar
    mensagens no Outlook classico. Permite configurar corpo em texto ou
    HTML, destinatarios adicionais, anexos, conta de envio, exibicao,
    salvamento e envio da resposta.
Dependencias:
    - Outlook classico para Windows
    - constantes.py
    - excecoes.py
============================================================
"""

from __future__ import annotations

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
        OL_TO,
    )
except ImportError:
    OL_FORMAT_PLAIN = 1
    OL_FORMAT_HTML = 2
    OL_TO = 1
    OL_CC = 2
    OL_BCC = 3


class OutlookRespostasError(AutomacaoError):
    """Erro base para respostas e encaminhamentos do Outlook."""


class MensagemOriginalInvalidaError(OutlookRespostasError):
    """A mensagem original esta ausente ou nao suporta a operacao."""


class RespostaCriacaoError(OutlookRespostasError):
    """O Outlook nao conseguiu criar a resposta ou encaminhamento."""


class RespostaConfiguracaoError(OutlookRespostasError):
    """A resposta nao pode ser configurada com os dados informados."""


class RespostaEnvioError(OutlookRespostasError):
    """A resposta ou encaminhamento nao pode ser enviado."""


class RespostaAcaoError(OutlookRespostasError):
    """Uma acao de salvar, exibir ou excluir a resposta falhou."""


class DestinatarioRespostaError(OutlookRespostasError):
    """Um destinatario adicional nao pode ser adicionado ou resolvido."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM e retorna um padrao em caso de erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise MensagemOriginalInvalidaError(
            f"O parametro '{nome}' nao pode ser None."
        )


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


def _validar_booleano(valor: bool, nome: str) -> None:
    """Valida uma opcao booleana."""
    if not isinstance(valor, bool):
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")


def _normalizar_destinatarios(
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
) -> list[str]:
    """Normaliza destinatarios e remove valores duplicados."""
    if isinstance(destinatarios, str):
        destinatarios = [
            item.strip()
            for item in destinatarios.replace(",", ";").split(";")
            if item.strip()
        ]
    if not isinstance(destinatarios, (list, tuple, set)) or not destinatarios:
        raise TypeError(
            "Destinatarios deve ser uma string ou colecao nao vazia."
        )
    resultado: list[str] = []
    for destinatario in destinatarios:
        texto = _validar_texto(destinatario, "destinatario")
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def _definir_propriedade(item: Any, propriedade: str, valor: Any) -> None:
    """Define uma propriedade COM com tratamento padronizado."""
    _validar_objeto(item, "item")
    try:
        setattr(item, propriedade, valor)
    except Exception as error:
        raise RespostaConfiguracaoError(
            f"Nao foi possivel definir a propriedade '{propriedade}'."
        ) from error


def criar_resposta(email_original: Any) -> Any:
    """Cria uma resposta destinada somente ao remetente original."""
    _validar_objeto(email_original, "email_original")
    try:
        resposta = email_original.Reply()
    except Exception as error:
        raise RespostaCriacaoError(
            "Nao foi possivel criar a resposta."
        ) from error
    if resposta is None:
        raise RespostaCriacaoError("O Outlook nao retornou um MailItem de resposta.")
    return resposta


def criar_resposta_todos(email_original: Any) -> Any:
    """Cria uma resposta para o remetente e destinatarios originais."""
    _validar_objeto(email_original, "email_original")
    try:
        resposta = email_original.ReplyAll()
    except Exception as error:
        raise RespostaCriacaoError(
            "Nao foi possivel criar a resposta para todos."
        ) from error
    if resposta is None:
        raise RespostaCriacaoError("O Outlook nao retornou um MailItem de resposta.")
    return resposta


def criar_encaminhamento(email_original: Any) -> Any:
    """Cria um encaminhamento da mensagem original."""
    _validar_objeto(email_original, "email_original")
    try:
        encaminhamento = email_original.Forward()
    except Exception as error:
        raise RespostaCriacaoError(
            "Nao foi possivel criar o encaminhamento."
        ) from error
    if encaminhamento is None:
        raise RespostaCriacaoError(
            "O Outlook nao retornou um MailItem de encaminhamento."
        )
    return encaminhamento


def criar_item_resposta(
    email_original: Any,
    responder_todos: bool = False,
    encaminhar: bool = False,
) -> Any:
    """Cria resposta simples, resposta a todos ou encaminhamento."""
    _validar_booleano(responder_todos, "responder_todos")
    _validar_booleano(encaminhar, "encaminhar")
    if responder_todos and encaminhar:
        raise ValueError(
            "Nao e permitido responder a todos e encaminhar simultaneamente."
        )
    if encaminhar:
        return criar_encaminhamento(email_original)
    if responder_todos:
        return criar_resposta_todos(email_original)
    return criar_resposta(email_original)


def definir_corpo_texto(
    resposta: Any,
    corpo: str,
    preservar_original: bool = True,
    inserir_antes: bool = True,
    separador: str = "\n\n",
) -> Any:
    """Define texto novo preservando opcionalmente o corpo original."""
    _validar_booleano(preservar_original, "preservar_original")
    _validar_booleano(inserir_antes, "inserir_antes")
    texto = _validar_texto(corpo, "corpo", permitir_vazio=True)
    divisor = _validar_texto(separador, "separador", permitir_vazio=True)
    original = str(_obter_propriedade(resposta, "Body", "") or "")
    if preservar_original:
        conteudo = (
            texto + divisor + original
            if inserir_antes
            else original + divisor + texto
        )
    else:
        conteudo = texto
    _definir_propriedade(resposta, "BodyFormat", OL_FORMAT_PLAIN)
    _definir_propriedade(resposta, "Body", conteudo)
    return resposta


def definir_corpo_html(
    resposta: Any,
    html: str,
    preservar_original: bool = True,
    inserir_antes: bool = True,
    separador: str = "<br><br>",
) -> Any:
    """Define HTML novo preservando opcionalmente o HTML original."""
    _validar_booleano(preservar_original, "preservar_original")
    _validar_booleano(inserir_antes, "inserir_antes")
    conteudo_novo = _validar_texto(html, "html", permitir_vazio=True)
    divisor = _validar_texto(separador, "separador", permitir_vazio=True)
    original = str(_obter_propriedade(resposta, "HTMLBody", "") or "")
    if preservar_original:
        conteudo = (
            conteudo_novo + divisor + original
            if inserir_antes
            else original + divisor + conteudo_novo
        )
    else:
        conteudo = conteudo_novo
    _definir_propriedade(resposta, "BodyFormat", OL_FORMAT_HTML)
    _definir_propriedade(resposta, "HTMLBody", conteudo)
    return resposta


def definir_assunto(resposta: Any, assunto: str) -> Any:
    """Substitui o assunto gerado automaticamente pelo Outlook."""
    _definir_propriedade(
        resposta,
        "Subject",
        _validar_texto(assunto, "assunto"),
    )
    return resposta


def adicionar_prefixo_assunto(
    resposta: Any,
    prefixo: str,
    evitar_duplicado: bool = True,
) -> Any:
    """Adiciona um prefixo ao assunto atual da resposta."""
    texto = _validar_texto(prefixo, "prefixo")
    _validar_booleano(evitar_duplicado, "evitar_duplicado")
    assunto = str(_obter_propriedade(resposta, "Subject", "") or "")
    if evitar_duplicado and assunto.lower().startswith(texto.lower()):
        return resposta
    _definir_propriedade(resposta, "Subject", f"{texto}{assunto}")
    return resposta


def adicionar_destinatario(
    resposta: Any,
    endereco: str,
    tipo: int = OL_TO,
) -> Any:
    """Adiciona um destinatario Para, CC ou CCO pela colecao Recipients."""
    _validar_objeto(resposta, "resposta")
    valor = _validar_texto(endereco, "endereco")
    if isinstance(tipo, bool) or not isinstance(tipo, int):
        raise TypeError("O parametro 'tipo' deve ser inteiro.")
    if tipo not in {OL_TO, OL_CC, OL_BCC}:
        raise ValueError("O tipo deve representar Para, CC ou CCO.")
    try:
        destinatario = resposta.Recipients.Add(valor)
        destinatario.Type = tipo
        return destinatario
    except Exception as error:
        raise DestinatarioRespostaError(
            f"Nao foi possivel adicionar o destinatario '{valor}'."
        ) from error


def adicionar_destinatarios(
    resposta: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    tipo: int = OL_TO,
) -> list[Any]:
    """Adiciona varios destinatarios do mesmo tipo."""
    return [
        adicionar_destinatario(resposta, endereco, tipo)
        for endereco in _normalizar_destinatarios(destinatarios)
    ]


def definir_destinatarios_encaminhamento(
    encaminhamento: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    copias: str | list[str] | tuple[str, ...] | set[str] | None = None,
    copias_ocultas: str | list[str] | tuple[str, ...] | set[str] | None = None,
) -> Any:
    """Define Para, CC e CCO em um encaminhamento."""
    _definir_propriedade(
        encaminhamento,
        "To",
        "; ".join(_normalizar_destinatarios(destinatarios)),
    )
    if copias is not None:
        _definir_propriedade(
            encaminhamento,
            "CC",
            "; ".join(_normalizar_destinatarios(copias)),
        )
    if copias_ocultas is not None:
        _definir_propriedade(
            encaminhamento,
            "BCC",
            "; ".join(_normalizar_destinatarios(copias_ocultas)),
        )
    return encaminhamento


def resolver_destinatarios(
    resposta: Any,
    exigir_resolucao: bool = True,
) -> bool:
    """Resolve todos os destinatarios da resposta."""
    _validar_objeto(resposta, "resposta")
    _validar_booleano(exigir_resolucao, "exigir_resolucao")
    try:
        resolvido = bool(resposta.Recipients.ResolveAll())
    except Exception as error:
        raise DestinatarioRespostaError(
            "Nao foi possivel resolver os destinatarios."
        ) from error
    if exigir_resolucao and not resolvido:
        raise DestinatarioRespostaError(
            "Um ou mais destinatarios nao foram resolvidos."
        )
    return resolvido


def remover_destinatario_por_indice(resposta: Any, indice: int) -> bool:
    """Remove um destinatario pelo indice COM iniciado em um."""
    _validar_objeto(resposta, "resposta")
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError("O parametro 'indice' deve ser inteiro.")
    if indice < 1:
        raise ValueError("O parametro 'indice' deve ser maior ou igual a 1.")
    try:
        resposta.Recipients.Remove(indice)
        return True
    except Exception as error:
        raise DestinatarioRespostaError(
            f"Nao foi possivel remover o destinatario {indice}."
        ) from error


def definir_conta_envio(resposta: Any, conta: Any) -> Any:
    """Define a conta utilizada para enviar a resposta."""
    _validar_objeto(resposta, "resposta")
    _validar_objeto(conta, "conta")
    _definir_propriedade(resposta, "SendUsingAccount", conta)
    return resposta


def adicionar_anexo(
    resposta: Any,
    caminho: str | Path,
    nome_exibicao: str | None = None,
) -> Any:
    """Adiciona um arquivo existente a resposta ou encaminhamento."""
    _validar_objeto(resposta, "resposta")
    if not isinstance(caminho, (str, Path)):
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    arquivo = Path(caminho).expanduser().resolve()
    if not arquivo.is_file():
        raise FileNotFoundError(f"O arquivo nao existe: {arquivo}")
    nome = (
        arquivo.name
        if nome_exibicao is None
        else _validar_texto(nome_exibicao, "nome_exibicao")
    )
    try:
        return resposta.Attachments.Add(str(arquivo), 1, 1, nome)
    except Exception as error:
        raise RespostaConfiguracaoError(
            f"Nao foi possivel adicionar o anexo '{arquivo}'."
        ) from error


def adicionar_anexos(
    resposta: Any,
    caminhos: list[str | Path] | tuple[str | Path, ...] | set[str | Path],
    ignorar_inexistentes: bool = False,
) -> list[Any]:
    """Adiciona varios arquivos a resposta."""
    if not isinstance(caminhos, (list, tuple, set)) or not caminhos:
        raise TypeError("O parametro 'caminhos' deve ser uma colecao nao vazia.")
    _validar_booleano(ignorar_inexistentes, "ignorar_inexistentes")
    anexos = []
    for caminho in caminhos:
        arquivo = Path(caminho).expanduser().resolve()
        if ignorar_inexistentes and not arquivo.is_file():
            continue
        anexos.append(adicionar_anexo(resposta, arquivo))
    return anexos


def configurar_resposta(
    resposta: Any,
    corpo: str,
    html: bool = False,
    preservar_original: bool = True,
    inserir_antes: bool = True,
    assunto: str | None = None,
    copias: str | list[str] | tuple[str, ...] | set[str] | None = None,
    copias_ocultas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    anexos: list[str | Path] | tuple[str | Path, ...] | set[str | Path] | None = None,
    conta_envio: Any | None = None,
) -> Any:
    """Configura corpo e opcoes comuns de uma resposta."""
    _validar_booleano(html, "html")
    if html:
        definir_corpo_html(
            resposta,
            corpo,
            preservar_original=preservar_original,
            inserir_antes=inserir_antes,
        )
    else:
        definir_corpo_texto(
            resposta,
            corpo,
            preservar_original=preservar_original,
            inserir_antes=inserir_antes,
        )
    if assunto is not None:
        definir_assunto(resposta, assunto)
    if copias is not None:
        adicionar_destinatarios(resposta, copias, OL_CC)
    if copias_ocultas is not None:
        adicionar_destinatarios(resposta, copias_ocultas, OL_BCC)
    if anexos is not None:
        adicionar_anexos(resposta, anexos)
    if conta_envio is not None:
        definir_conta_envio(resposta, conta_envio)
    return resposta


def configurar_encaminhamento(
    encaminhamento: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    corpo: str = "",
    html: bool = False,
    preservar_original: bool = True,
    copias: str | list[str] | tuple[str, ...] | set[str] | None = None,
    copias_ocultas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    anexos: list[str | Path] | tuple[str | Path, ...] | set[str | Path] | None = None,
    conta_envio: Any | None = None,
) -> Any:
    """Configura destinatarios e conteudo de um encaminhamento."""
    definir_destinatarios_encaminhamento(
        encaminhamento,
        destinatarios,
        copias=copias,
        copias_ocultas=copias_ocultas,
    )
    configurar_resposta(
        encaminhamento,
        corpo,
        html=html,
        preservar_original=preservar_original,
        inserir_antes=True,
        anexos=anexos,
        conta_envio=conta_envio,
    )
    return encaminhamento


def salvar_resposta(resposta: Any) -> Any:
    """Salva a resposta como rascunho."""
    _validar_objeto(resposta, "resposta")
    try:
        resposta.Save()
        return resposta
    except Exception as error:
        raise RespostaAcaoError(
            "Nao foi possivel salvar a resposta."
        ) from error


def exibir_resposta(resposta: Any, modal: bool = False) -> Any:
    """Exibe a resposta ou encaminhamento no Outlook."""
    _validar_objeto(resposta, "resposta")
    _validar_booleano(modal, "modal")
    try:
        resposta.Display(modal)
        return resposta
    except Exception as error:
        raise RespostaAcaoError(
            "Nao foi possivel exibir a resposta."
        ) from error


def enviar_resposta(resposta: Any, resolver: bool = True) -> Any:
    """Resolve destinatarios e envia a resposta."""
    _validar_objeto(resposta, "resposta")
    _validar_booleano(resolver, "resolver")
    if resolver:
        resolver_destinatarios(resposta, exigir_resolucao=True)
    try:
        resposta.Send()
        return resposta
    except Exception as error:
        raise RespostaEnvioError(
            "Nao foi possivel enviar a resposta."
        ) from error


def excluir_resposta(resposta: Any) -> bool:
    """Exclui uma resposta salva ou em rascunho."""
    _validar_objeto(resposta, "resposta")
    try:
        resposta.Delete()
        return True
    except Exception as error:
        raise RespostaAcaoError(
            "Nao foi possivel excluir a resposta."
        ) from error


def responder_email(
    email_original: Any,
    corpo: str,
    responder_todos: bool = False,
    html: bool = False,
    enviar: bool = False,
    exibir: bool = False,
    salvar: bool = False,
    **configuracoes: Any,
) -> Any:
    """Cria e processa uma resposta simples ou para todos."""
    for nome, valor in {
        "responder_todos": responder_todos,
        "html": html,
        "enviar": enviar,
        "exibir": exibir,
        "salvar": salvar,
    }.items():
        _validar_booleano(valor, nome)
    resposta = criar_item_resposta(
        email_original,
        responder_todos=responder_todos,
    )
    configurar_resposta(resposta, corpo, html=html, **configuracoes)
    if salvar:
        salvar_resposta(resposta)
    if exibir:
        exibir_resposta(resposta)
    if enviar:
        enviar_resposta(resposta)
    return resposta


def encaminhar_email(
    email_original: Any,
    destinatarios: str | list[str] | tuple[str, ...] | set[str],
    corpo: str = "",
    html: bool = False,
    enviar: bool = False,
    exibir: bool = False,
    salvar: bool = False,
    **configuracoes: Any,
) -> Any:
    """Cria, configura e processa um encaminhamento."""
    for nome, valor in {
        "html": html,
        "enviar": enviar,
        "exibir": exibir,
        "salvar": salvar,
    }.items():
        _validar_booleano(valor, nome)
    encaminhamento = criar_encaminhamento(email_original)
    configurar_encaminhamento(
        encaminhamento,
        destinatarios,
        corpo=corpo,
        html=html,
        **configuracoes,
    )
    if salvar:
        salvar_resposta(encaminhamento)
    if exibir:
        exibir_resposta(encaminhamento)
    if enviar:
        enviar_resposta(encaminhamento)
    return encaminhamento


def obter_dados_resposta(resposta: Any) -> dict[str, Any]:
    """Converte as principais propriedades da resposta para dicionario."""
    _validar_objeto(resposta, "resposta")
    anexos = _obter_propriedade(resposta, "Attachments")
    destinatarios = _obter_propriedade(resposta, "Recipients")
    return {
        "entry_id": _obter_propriedade(resposta, "EntryID"),
        "assunto": str(_obter_propriedade(resposta, "Subject", "") or ""),
        "para": str(_obter_propriedade(resposta, "To", "") or ""),
        "cc": str(_obter_propriedade(resposta, "CC", "") or ""),
        "cco": str(_obter_propriedade(resposta, "BCC", "") or ""),
        "corpo": str(_obter_propriedade(resposta, "Body", "") or ""),
        "corpo_html": str(_obter_propriedade(resposta, "HTMLBody", "") or ""),
        "formato_corpo": _obter_propriedade(resposta, "BodyFormat"),
        "salva": bool(_obter_propriedade(resposta, "Saved", False)),
        "quantidade_destinatarios": int(
            _obter_propriedade(destinatarios, "Count", 0) or 0
        ),
        "quantidade_anexos": int(
            _obter_propriedade(anexos, "Count", 0) or 0
        ),
        "conta_envio": _obter_propriedade(resposta, "SendUsingAccount"),
    }


__all__ = [
    "OutlookRespostasError",
    "MensagemOriginalInvalidaError",
    "RespostaCriacaoError",
    "RespostaConfiguracaoError",
    "RespostaEnvioError",
    "RespostaAcaoError",
    "DestinatarioRespostaError",
    "criar_resposta",
    "criar_resposta_todos",
    "criar_encaminhamento",
    "criar_item_resposta",
    "definir_corpo_texto",
    "definir_corpo_html",
    "definir_assunto",
    "adicionar_prefixo_assunto",
    "adicionar_destinatario",
    "adicionar_destinatarios",
    "definir_destinatarios_encaminhamento",
    "resolver_destinatarios",
    "remover_destinatario_por_indice",
    "definir_conta_envio",
    "adicionar_anexo",
    "adicionar_anexos",
    "configurar_resposta",
    "configurar_encaminhamento",
    "salvar_resposta",
    "exibir_resposta",
    "enviar_resposta",
    "excluir_resposta",
    "responder_email",
    "encaminhar_email",
    "obter_dados_resposta",
]
