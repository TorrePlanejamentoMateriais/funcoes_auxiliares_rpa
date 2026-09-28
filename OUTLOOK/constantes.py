"""
============================================================
Modulo: OUTLOOK / constantes.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Centraliza constantes e enumeracoes utilizadas na automacao do
    Outlook classico via COM. O modulo evita numeros magicos e permite
    utilizar valores tipados ou aliases compativeis com codigo legado.
Dependencias:
    - Nenhuma dependencia externa
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de itens, pastas, anexos, destinatarios e reunioes.
        - Inclusao de importancia, formato, resposta e recorrencia.
        - Inclusao de aliases e funcoes auxiliares de consulta.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa IntEnum e IntFlag para enumeracoes compativeis com inteiros COM
from enum import IntEnum, IntFlag

# Importa TypeVar para tipar funcoes auxiliares genericas
from typing import TypeVar


# ----------------------------------------------------------------------------
# CONSTANTES GERAIS
# ----------------------------------------------------------------------------

# ProgID utilizado pelo pywin32 para obter Outlook.Application
OUTLOOK_PROG_ID = "Outlook.Application"

# Nome do unico Namespace suportado pelo modelo de objetos do Outlook
NAMESPACE_MAPI = "MAPI"

# Nome da classe de mensagem padrao de e-mail
MESSAGE_CLASS_MAIL = "IPM.Note"

# Nome da classe de mensagem padrao de compromisso
MESSAGE_CLASS_APPOINTMENT = "IPM.Appointment"

# Nome da classe de mensagem padrao de contato
MESSAGE_CLASS_CONTACT = "IPM.Contact"

# Nome da classe de mensagem padrao de tarefa
MESSAGE_CLASS_TASK = "IPM.Task"

# Separador comumente utilizado pelo Outlook para categorias
SEPARADOR_CATEGORIAS = ", "

# Formato normalmente utilizado em filtros Restrict de data
FORMATO_DATA_OUTLOOK = "%m/%d/%Y %I:%M %p"


# ----------------------------------------------------------------------------
# TIPOS DE ITEM
# ----------------------------------------------------------------------------

class TipoItemOutlook(IntEnum):
    """Valores da enumeracao OlItemType do Outlook."""

    EMAIL = 0
    COMPROMISSO = 1
    CONTATO = 2
    TAREFA = 3
    DIARIO = 4
    NOTA = 5
    POSTAGEM = 6
    LISTA_DISTRIBUICAO = 7
    COMPARTILHAMENTO = 104
    DOCUMENTO = 41


# ----------------------------------------------------------------------------
# PASTAS PADRAO
# ----------------------------------------------------------------------------

class PastaPadraoOutlook(IntEnum):
    """Valores da enumeracao OlDefaultFolders do Outlook."""

    ITENS_EXCLUIDOS = 3
    CAIXA_SAIDA = 4
    ITENS_ENVIADOS = 5
    CAIXA_ENTRADA = 6
    CALENDARIO = 9
    CONTATOS = 10
    DIARIO = 11
    ANOTACOES = 12
    TAREFAS = 13
    RASCUNHOS = 16
    PASTAS_PUBLICAS = 18
    CONFLITOS = 19
    PROBLEMAS_SINCRONIZACAO = 20
    FALHAS_LOCAIS = 21
    FALHAS_SERVIDOR = 22
    LIXO_ELETRONICO = 23
    RSS_FEEDS = 25
    TAREFAS_PENDENTES = 28
    EMAIL_GERENCIADO = 29
    CONTATOS_SUGERIDOS = 30


# ----------------------------------------------------------------------------
# ANEXOS
# ----------------------------------------------------------------------------

class TipoAnexo(IntEnum):
    """Valores da enumeracao OlAttachmentType do Outlook."""

    POR_VALOR = 1
    POR_REFERENCIA = 4
    INCORPORADO = 5
    OLE = 6


# ----------------------------------------------------------------------------
# IMPORTANCIA, SENSIBILIDADE E FORMATO
# ----------------------------------------------------------------------------

class Importancia(IntEnum):
    """Nivel de importancia de itens do Outlook."""

    BAIXA = 0
    NORMAL = 1
    ALTA = 2


class Sensibilidade(IntEnum):
    """Nivel de sensibilidade de itens do Outlook."""

    NORMAL = 0
    PESSOAL = 1
    PRIVADO = 2
    CONFIDENCIAL = 3


class FormatoCorpo(IntEnum):
    """Formato do corpo de uma mensagem Outlook."""

    NAO_ESPECIFICADO = 0
    TEXTO = 1
    HTML = 2
    RTF = 3


class FormatoEditor(IntEnum):
    """Formato usado por editores e inspetores do Outlook."""

    TEXTO = 1
    HTML = 2
    RTF = 3


# ----------------------------------------------------------------------------
# DESTINATARIOS E CONTAS
# ----------------------------------------------------------------------------

class TipoDestinatario(IntEnum):
    """Tipos de destinatario de mensagens, contatos e reunioes."""

    ORGANIZADOR = 0
    OBRIGATORIO = 1
    OPCIONAL = 2
    RECURSO = 3

    # Em MailItem os mesmos valores representam Para, CC e CCO
    PARA = 1
    CC = 2
    CCO = 3


class TipoConta(IntEnum):
    """Valores comuns da enumeracao OlAccountType."""

    EXCHANGE = 0
    IMAP = 1
    POP3 = 2
    HTTP = 3
    OUTRO = 5
    EAS = 6


class TipoEndereco(IntEnum):
    """Tipos comuns de AddressEntry do Outlook."""

    OUTLOOK = 0
    EXCHANGE = 0
    SMTP = 1


# ----------------------------------------------------------------------------
# COMPROMISSOS E REUNIOES
# ----------------------------------------------------------------------------

class StatusReuniao(IntEnum):
    """Estado de um AppointmentItem em relacao a reunioes."""

    NAO_REUNIAO = 0
    REUNIAO = 1
    RECEBIDA = 3
    CANCELADA = 5
    CANCELAMENTO_RECEBIDO = 7


class OcupacaoCalendario(IntEnum):
    """Estado de disponibilidade exibido no calendario."""

    LIVRE = 0
    PROVISORIO = 1
    OCUPADO = 2
    FORA_ESCRITORIO = 3
    TRABALHANDO_OUTRO_LOCAL = 4


class RespostaReuniao(IntEnum):
    """Estado ou acao de resposta a solicitacao de reuniao."""

    NENHUMA = 0
    ORGANIZADA = 1
    PROVISORIA = 2
    ACEITA = 3
    RECUSADA = 4
    NAO_RESPONDIDA = 5


class AcaoRespostaReuniao(IntEnum):
    """Acao usada em AppointmentItem.Respond."""

    PROVISORIA = 2
    ACEITAR = 3
    RECUSAR = 4


class ModoRespostaReuniao(IntEnum):
    """Modo de envio da resposta de reuniao."""

    SEM_RESPOSTA = 0
    ORGANIZADOR = 1
    TODOS = 2


# ----------------------------------------------------------------------------
# RECORRENCIA E DIAS DA SEMANA
# ----------------------------------------------------------------------------

class TipoRecorrencia(IntEnum):
    """Tipos de recorrencia suportados por RecurrencePattern."""

    DIARIA = 0
    SEMANAL = 1
    MENSAL = 2
    MENSAL_POSICAO = 3
    ANUAL = 5
    ANUAL_POSICAO = 6


class DiaSemana(IntFlag):
    """Mascara de dias utilizada por RecurrencePattern.DayOfWeekMask."""

    DOMINGO = 1
    SEGUNDA = 2
    TERCA = 4
    QUARTA = 8
    QUINTA = 16
    SEXTA = 32
    SABADO = 64

    DIAS_UTEIS = SEGUNDA | TERCA | QUARTA | QUINTA | SEXTA
    FIM_DE_SEMANA = DOMINGO | SABADO
    TODOS = DOMINGO | SEGUNDA | TERCA | QUARTA | QUINTA | SEXTA | SABADO


class InstanciaRecorrencia(IntEnum):
    """Posicao ordinal usada em recorrencias mensais ou anuais."""

    PRIMEIRA = 1
    SEGUNDA = 2
    TERCEIRA = 3
    QUARTA = 4
    ULTIMA = 5


# ----------------------------------------------------------------------------
# TAREFAS E MARCACOES
# ----------------------------------------------------------------------------

class StatusTarefa(IntEnum):
    """Status de um TaskItem do Outlook."""

    NAO_INICIADA = 0
    EM_ANDAMENTO = 1
    CONCLUIDA = 2
    AGUARDANDO_OUTRA_PESSOA = 3
    ADIADA = 4


class StatusMarcacao(IntEnum):
    """Estado de marcacao para acompanhamento."""

    NAO_MARCADO = 0
    CONCLUIDO = 1
    MARCADO = 2


# ----------------------------------------------------------------------------
# ACOES DE EXIBICAO, SALVAMENTO E FECHAMENTO
# ----------------------------------------------------------------------------

class ModoJanela(IntEnum):
    """Modo de exibicao das janelas do Outlook."""

    NORMAL = 0
    MINIMIZADA = 1
    MAXIMIZADA = 2


class ModoFechamento(IntEnum):
    """Acao executada ao fechar um Inspector ou item."""

    SALVAR = 0
    DESCARTAR = 1
    PERGUNTAR = 2


class FormatoSalvamento(IntEnum):
    """Formatos utilizados pelo metodo SaveAs de itens Outlook."""

    TEXTO = 0
    RTF = 1
    MODELO = 2
    MENSAGEM = 3
    ICAL = 8
    VCARD = 6
    HTML = 5
    MHTML = 10
    MENSAGEM_UNICODE = 9


# ----------------------------------------------------------------------------
# REGRAS, EVENTOS E PESQUISA
# ----------------------------------------------------------------------------

class TipoRegra(IntEnum):
    """Tipo de regra criada na colecao Rules."""

    RECEBIMENTO = 0
    ENVIO = 1


class EscopoPesquisa(IntEnum):
    """Escopo usado em pesquisas avancadas do Outlook."""

    PASTA_ATUAL = 0
    SUBPASTAS = 1
    TODAS_PASTAS = 2


class OrdemClassificacao(IntEnum):
    """Direcao de ordenacao usada por colecoes e tabelas."""

    CRESCENTE = 1
    DECRESCENTE = 2


# ----------------------------------------------------------------------------
# ALIASES COMPATIVEIS COM CODIGO EXISTENTE
# ----------------------------------------------------------------------------

# Tipos de item
OL_MAIL_ITEM = int(TipoItemOutlook.EMAIL)
OL_APPOINTMENT_ITEM = int(TipoItemOutlook.COMPROMISSO)
OL_CONTACT_ITEM = int(TipoItemOutlook.CONTATO)
OL_TASK_ITEM = int(TipoItemOutlook.TAREFA)
OL_JOURNAL_ITEM = int(TipoItemOutlook.DIARIO)
OL_NOTE_ITEM = int(TipoItemOutlook.NOTA)
OL_POST_ITEM = int(TipoItemOutlook.POSTAGEM)
OL_DISTRIBUTION_LIST_ITEM = int(TipoItemOutlook.LISTA_DISTRIBUICAO)

# Pastas padrao
OL_FOLDER_DELETED_ITEMS = int(PastaPadraoOutlook.ITENS_EXCLUIDOS)
OL_FOLDER_OUTBOX = int(PastaPadraoOutlook.CAIXA_SAIDA)
OL_FOLDER_SENT_MAIL = int(PastaPadraoOutlook.ITENS_ENVIADOS)
OL_FOLDER_INBOX = int(PastaPadraoOutlook.CAIXA_ENTRADA)
OL_FOLDER_CALENDAR = int(PastaPadraoOutlook.CALENDARIO)
OL_FOLDER_CONTACTS = int(PastaPadraoOutlook.CONTATOS)
OL_FOLDER_JOURNAL = int(PastaPadraoOutlook.DIARIO)
OL_FOLDER_NOTES = int(PastaPadraoOutlook.ANOTACOES)
OL_FOLDER_TASKS = int(PastaPadraoOutlook.TAREFAS)
OL_FOLDER_DRAFTS = int(PastaPadraoOutlook.RASCUNHOS)
OL_FOLDER_JUNK = int(PastaPadraoOutlook.LIXO_ELETRONICO)
OL_FOLDER_RSS_FEEDS = int(PastaPadraoOutlook.RSS_FEEDS)
OL_FOLDER_TO_DO = int(PastaPadraoOutlook.TAREFAS_PENDENTES)

# Anexos
OL_BY_VALUE = int(TipoAnexo.POR_VALOR)
OL_BY_REFERENCE = int(TipoAnexo.POR_REFERENCIA)
OL_EMBEDDED_ITEM = int(TipoAnexo.INCORPORADO)
OL_OLE = int(TipoAnexo.OLE)

# Importancia e sensibilidade
OL_IMPORTANCE_LOW = int(Importancia.BAIXA)
OL_IMPORTANCE_NORMAL = int(Importancia.NORMAL)
OL_IMPORTANCE_HIGH = int(Importancia.ALTA)
OL_NORMAL = int(Sensibilidade.NORMAL)
OL_PERSONAL = int(Sensibilidade.PESSOAL)
OL_PRIVATE = int(Sensibilidade.PRIVADO)
OL_CONFIDENTIAL = int(Sensibilidade.CONFIDENCIAL)

# Corpo da mensagem
OL_FORMAT_UNSPECIFIED = int(FormatoCorpo.NAO_ESPECIFICADO)
OL_FORMAT_PLAIN = int(FormatoCorpo.TEXTO)
OL_FORMAT_HTML = int(FormatoCorpo.HTML)
OL_FORMAT_RICH_TEXT = int(FormatoCorpo.RTF)

# Destinatarios
OL_ORGANIZER = int(TipoDestinatario.ORGANIZADOR)
OL_REQUIRED = int(TipoDestinatario.OBRIGATORIO)
OL_OPTIONAL = int(TipoDestinatario.OPCIONAL)
OL_RESOURCE = int(TipoDestinatario.RECURSO)
OL_TO = int(TipoDestinatario.PARA)
OL_CC = int(TipoDestinatario.CC)
OL_BCC = int(TipoDestinatario.CCO)

# Reunioes e calendario
OL_NON_MEETING = int(StatusReuniao.NAO_REUNIAO)
OL_MEETING = int(StatusReuniao.REUNIAO)
OL_MEETING_RECEIVED = int(StatusReuniao.RECEBIDA)
OL_MEETING_CANCELED = int(StatusReuniao.CANCELADA)
OL_FREE = int(OcupacaoCalendario.LIVRE)
OL_TENTATIVE = int(OcupacaoCalendario.PROVISORIO)
OL_BUSY = int(OcupacaoCalendario.OCUPADO)
OL_OUT_OF_OFFICE = int(OcupacaoCalendario.FORA_ESCRITORIO)
OL_WORKING_ELSEWHERE = int(OcupacaoCalendario.TRABALHANDO_OUTRO_LOCAL)
OL_MEETING_ACCEPT = int(AcaoRespostaReuniao.ACEITAR)
OL_MEETING_DECLINE = int(AcaoRespostaReuniao.RECUSAR)
OL_MEETING_TENTATIVE = int(AcaoRespostaReuniao.PROVISORIA)

# Recorrencia
OL_RECURRENCE_DAILY = int(TipoRecorrencia.DIARIA)
OL_RECURRENCE_WEEKLY = int(TipoRecorrencia.SEMANAL)
OL_RECURRENCE_MONTHLY = int(TipoRecorrencia.MENSAL)
OL_RECURRENCE_MONTH_NTH = int(TipoRecorrencia.MENSAL_POSICAO)
OL_RECURRENCE_YEARLY = int(TipoRecorrencia.ANUAL)
OL_RECURRENCE_YEAR_NTH = int(TipoRecorrencia.ANUAL_POSICAO)

# Modos de fechamento
OL_SAVE = int(ModoFechamento.SALVAR)
OL_DISCARD = int(ModoFechamento.DESCARTAR)
OL_PROMPT_FOR_SAVE = int(ModoFechamento.PERGUNTAR)


# ----------------------------------------------------------------------------
# MAPAS DE APOIO
# ----------------------------------------------------------------------------

# Permite localizar pastas padrao por nomes amigaveis
MAPA_PASTAS_PADRAO = {
    "itens_excluidos": PastaPadraoOutlook.ITENS_EXCLUIDOS,
    "caixa_saida": PastaPadraoOutlook.CAIXA_SAIDA,
    "itens_enviados": PastaPadraoOutlook.ITENS_ENVIADOS,
    "caixa_entrada": PastaPadraoOutlook.CAIXA_ENTRADA,
    "calendario": PastaPadraoOutlook.CALENDARIO,
    "contatos": PastaPadraoOutlook.CONTATOS,
    "diario": PastaPadraoOutlook.DIARIO,
    "anotacoes": PastaPadraoOutlook.ANOTACOES,
    "tarefas": PastaPadraoOutlook.TAREFAS,
    "rascunhos": PastaPadraoOutlook.RASCUNHOS,
    "lixo_eletronico": PastaPadraoOutlook.LIXO_ELETRONICO,
    "rss_feeds": PastaPadraoOutlook.RSS_FEEDS,
    "tarefas_pendentes": PastaPadraoOutlook.TAREFAS_PENDENTES,
}

# Descricoes amigaveis dos principais tipos de item
DESCRICOES_TIPO_ITEM = {
    TipoItemOutlook.EMAIL: "E-mail",
    TipoItemOutlook.COMPROMISSO: "Compromisso",
    TipoItemOutlook.CONTATO: "Contato",
    TipoItemOutlook.TAREFA: "Tarefa",
    TipoItemOutlook.DIARIO: "Diario",
    TipoItemOutlook.NOTA: "Anotacao",
    TipoItemOutlook.POSTAGEM: "Postagem",
    TipoItemOutlook.LISTA_DISTRIBUICAO: "Lista de distribuicao",
}

# Descricoes amigaveis dos estados de disponibilidade
DESCRICOES_OCUPACAO = {
    OcupacaoCalendario.LIVRE: "Livre",
    OcupacaoCalendario.PROVISORIO: "Provisorio",
    OcupacaoCalendario.OCUPADO: "Ocupado",
    OcupacaoCalendario.FORA_ESCRITORIO: "Fora do escritorio",
    OcupacaoCalendario.TRABALHANDO_OUTRO_LOCAL: "Trabalhando em outro local",
}


# ----------------------------------------------------------------------------
# FUNCOES AUXILIARES
# ----------------------------------------------------------------------------

# Tipo generico para classes derivadas de IntEnum
TEnum = TypeVar("TEnum", bound=IntEnum)


def obter_valor_constante(valor: int | IntEnum) -> int:
    """
    Converte uma constante inteira ou IntEnum para int.

    Args:
        valor (int | IntEnum):
            Valor que sera enviado para uma interface COM.

    Returns:
        int:
            Representacao inteira da constante.

    Raises:
        TypeError:
            Caso o valor nao seja int nem IntEnum.
    """
    # Rejeita booleanos porque bool herda de int
    if isinstance(valor, bool) or not isinstance(valor, (int, IntEnum)):
        raise TypeError("O valor deve ser um numero inteiro ou IntEnum.")

    return int(valor)


def obter_enum_por_valor(
    classe_enum: type[TEnum],
    valor: int,
) -> TEnum:
    """
    Converte um inteiro para uma enumeracao especifica.

    Args:
        classe_enum (type[TEnum]):
            Classe derivada de IntEnum.
        valor (int):
            Valor numerico da constante.

    Returns:
        TEnum:
            Membro correspondente da enumeracao.

    Raises:
        TypeError:
            Caso a classe ou o valor possuam tipos invalidos.
        ValueError:
            Caso o valor nao exista na enumeracao.
    """
    # Valida se a classe recebida representa uma enumeracao inteira
    if not isinstance(classe_enum, type) or not issubclass(classe_enum, IntEnum):
        raise TypeError("O parametro 'classe_enum' deve derivar de IntEnum.")

    # Rejeita booleanos e valores nao inteiros
    if isinstance(valor, bool) or not isinstance(valor, int):
        raise TypeError("O parametro 'valor' deve ser um numero inteiro.")

    return classe_enum(valor)


def obter_pasta_padrao(nome: str) -> PastaPadraoOutlook:
    """
    Retorna uma pasta padrao a partir de um nome amigavel.

    Args:
        nome (str):
            Nome como caixa_entrada, calendario ou rascunhos.

    Returns:
        PastaPadraoOutlook:
            Constante da pasta solicitada.

    Raises:
        TypeError:
            Caso o nome nao seja uma string.
        ValueError:
            Caso a pasta nao esteja cadastrada.
    """
    # Valida e normaliza o nome recebido
    if not isinstance(nome, str):
        raise TypeError("O parametro 'nome' deve ser uma string.")

    chave = nome.strip().lower().replace(" ", "_")
    if not chave:
        raise ValueError("O parametro 'nome' nao pode estar vazio.")

    try:
        return MAPA_PASTAS_PADRAO[chave]
    except KeyError as error:
        opcoes = ", ".join(sorted(MAPA_PASTAS_PADRAO))
        raise ValueError(
            f"Pasta padrao desconhecida: '{nome}'. Opcoes: {opcoes}."
        ) from error


def combinar_dias_semana(
    *dias: DiaSemana | int,
) -> DiaSemana:
    """
    Combina dias em uma mascara para recorrencias semanais.

    Args:
        *dias (DiaSemana | int):
            Dias ou mascaras que serao combinados.

    Returns:
        DiaSemana:
            Mascara resultante da operacao OR.

    Raises:
        ValueError:
            Caso nenhum dia seja informado.
        TypeError:
            Caso um valor nao seja inteiro ou DiaSemana.
    """
    # Exige ao menos um valor
    if not dias:
        raise ValueError("Informe ao menos um dia da semana.")

    resultado = DiaSemana(0)
    for dia in dias:
        if isinstance(dia, bool) or not isinstance(dia, (int, DiaSemana)):
            raise TypeError("Todos os dias devem ser inteiros ou DiaSemana.")
        resultado |= DiaSemana(int(dia))

    return resultado


def listar_constantes(
    classe_enum: type[TEnum],
) -> dict[str, int]:
    """
    Retorna os nomes e valores de uma enumeracao Outlook.

    Args:
        classe_enum (type[TEnum]):
            Classe derivada de IntEnum.

    Returns:
        dict[str, int]:
            Nomes dos membros e respectivos valores numericos.
    """
    # Valida a classe antes de percorrer seus membros
    if not isinstance(classe_enum, type) or not issubclass(classe_enum, IntEnum):
        raise TypeError("O parametro 'classe_enum' deve derivar de IntEnum.")

    return {
        membro.name: int(membro.value)
        for membro in classe_enum
    }
