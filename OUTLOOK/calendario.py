"""
============================================================
Modulo: OUTLOOK / calendario.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para criar, configurar, consultar,
    salvar, enviar, responder e cancelar compromissos e reunioes no
    Outlook classico utilizando automacao COM.
Funcoes disponiveis:
    - obter_pasta_calendario()
    - criar_compromisso()
    - configurar_compromisso()
    - criar_compromisso_completo()
    - criar_reuniao()
    - definir_assunto()
    - definir_corpo()
    - definir_local()
    - definir_inicio_fim()
    - definir_duracao()
    - definir_dia_inteiro()
    - definir_lembrete()
    - definir_ocupacao()
    - definir_importancia()
    - definir_categorias()
    - adicionar_participante()
    - adicionar_participantes()
    - resolver_participantes()
    - listar_participantes()
    - remover_participante()
    - remover_todos_participantes()
    - salvar_compromisso()
    - exibir_compromisso()
    - enviar_convite()
    - cancelar_compromisso()
    - excluir_compromisso()
    - listar_compromissos()
    - buscar_compromissos()
    - obter_compromisso_por_entry_id()
    - obter_dados_compromisso()
    - compromisso_possui_conflito()
    - definir_recorrencia_diaria()
    - definir_recorrencia_semanal()
    - definir_recorrencia_mensal()
    - remover_recorrencia()
Dependencias:
    - Outlook classico para Windows
    - pywin32 durante a execucao real
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de compromissos e reunioes.
        - Inclusao de participantes, lembretes e recorrencia.
        - Inclusao de pesquisa e extracao de dados do calendario.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa date, datetime e timedelta para controlar datas do calendario
from datetime import date, datetime, time, timedelta

# Importa Any para objetos COM dinamicos do Outlook
from typing import Any

# Importa a excecao principal da biblioteca quando disponivel
try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base utilizada quando o modulo central nao esta disponivel."""


# ----------------------------------------------------------------------------
# CONSTANTES DO OUTLOOK
# ----------------------------------------------------------------------------

# Tipo de item criado pelo Outlook para compromissos
OL_APPOINTMENT_ITEM = 1

# Pasta padrao de calendario
OL_FOLDER_CALENDAR = 9

# Estado do compromisso
OL_NON_MEETING = 0
OL_MEETING = 1
OL_MEETING_RECEIVED = 3
OL_MEETING_CANCELED = 5

# Tipos de participante
OL_ORGANIZER = 0
OL_REQUIRED = 1
OL_OPTIONAL = 2
OL_RESOURCE = 3

# Estado de disponibilidade
OL_FREE = 0
OL_TENTATIVE = 1
OL_BUSY = 2
OL_OUT_OF_OFFICE = 3
OL_WORKING_ELSEWHERE = 4

# Nivel de importancia
OL_IMPORTANCE_LOW = 0
OL_IMPORTANCE_NORMAL = 1
OL_IMPORTANCE_HIGH = 2

# Tipos de recorrencia
OL_RECURRENCE_DAILY = 0
OL_RECURRENCE_WEEKLY = 1
OL_RECURRENCE_MONTHLY = 2
OL_RECURRENCE_MONTH_NTH = 3

# Respostas para solicitacao de reuniao
OL_MEETING_ACCEPT = 3
OL_MEETING_DECLINE = 4
OL_MEETING_TENTATIVE = 2

# Formato padrao utilizado em filtros Restrict do Outlook
FORMATO_DATA_OUTLOOK = "%m/%d/%Y %I:%M %p"


# ----------------------------------------------------------------------------
# EXCECOES DO MODULO
# ----------------------------------------------------------------------------

class CalendarioOutlookError(AutomacaoError):
    """Erro base para operacoes relacionadas ao calendario do Outlook."""


class OutlookInvalidoError(CalendarioOutlookError):
    """O objeto Outlook, Namespace ou item informado nao e valido."""


class CompromissoCriacaoError(CalendarioOutlookError):
    """O compromisso nao pode ser criado pelo Outlook."""


class CompromissoConfiguracaoError(CalendarioOutlookError):
    """Uma propriedade do compromisso nao pode ser configurada."""


class CompromissoNaoEncontradoError(CalendarioOutlookError):
    """Nenhum compromisso correspondente foi encontrado."""


class ParticipanteInvalidoError(CalendarioOutlookError):
    """Um participante nao pode ser adicionado ou resolvido."""


class CompromissoEnvioError(CalendarioOutlookError):
    """O convite de reuniao nao pode ser enviado."""


class RecorrenciaError(CalendarioOutlookError):
    """A recorrencia do compromisso nao pode ser configurada."""


# ----------------------------------------------------------------------------
# FUNCOES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------

def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """
    Obtem uma propriedade de um objeto COM de forma segura.

    Args:
        objeto (Any):
            Objeto COM cuja propriedade sera consultada.
        nome (str):
            Nome da propriedade.
        padrao (Any):
            Valor retornado quando a consulta falhar.

    Returns:
        Any:
            Valor da propriedade ou o valor padrao.
    """
    # Tenta acessar a propriedade dinamica do objeto COM
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


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
    # Rejeita objetos ausentes antes das chamadas COM
    if objeto is None:
        raise OutlookInvalidoError(
            f"O parametro '{nome}' nao pode ser None."
        )


def _validar_texto(valor: str, nome: str, permitir_vazio: bool = False) -> str:
    """
    Valida e normaliza uma propriedade textual.

    Args:
        valor (str):
            Texto que sera validado.
        nome (str):
            Nome utilizado nas mensagens de erro.
        permitir_vazio (bool):
            Define se uma string vazia sera aceita.

    Returns:
        str:
            Texto normalizado.
    """
    # Valida a opcao booleana
    if not isinstance(permitir_vazio, bool):
        raise TypeError("O parametro 'permitir_vazio' deve ser booleano.")

    # Exige uma string
    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")

    # Remove espacos externos
    resultado = valor.strip()

    # Rejeita texto vazio quando ele e obrigatorio
    if not resultado and not permitir_vazio:
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")

    return resultado


def _converter_datetime(valor: datetime | date, nome: str) -> datetime:
    """
    Converte date ou datetime para datetime.

    Args:
        valor (datetime | date):
            Data que sera convertida.
        nome (str):
            Nome utilizado na mensagem de erro.

    Returns:
        datetime:
            Data e hora normalizadas.
    """
    # Mantem datetime porque ele tambem e uma subclasse de date
    if isinstance(valor, datetime):
        return valor

    # Converte date utilizando meia-noite
    if isinstance(valor, date):
        return datetime.combine(valor, time.min)

    raise TypeError(
        f"O parametro '{nome}' deve ser date ou datetime."
    )


def _validar_inteiro(
    valor: int,
    nome: str,
    minimo: int | None = None,
) -> None:
    """
    Valida um parametro inteiro e seu limite minimo.

    Args:
        valor (int):
            Valor que sera validado.
        nome (str):
            Nome utilizado na mensagem de erro.
        minimo (int | None):
            Menor valor permitido.

    Returns:
        None:
            A funcao apenas realiza a validacao.
    """
    # Rejeita booleanos e valores nao inteiros
    if isinstance(valor, bool) or not isinstance(valor, int):
        raise TypeError(f"O parametro '{nome}' deve ser inteiro.")

    # Valida o limite quando informado
    if minimo is not None and valor < minimo:
        raise ValueError(
            f"O parametro '{nome}' deve ser maior ou igual a {minimo}."
        )


def _normalizar_lista_textos(
    valores: str | list[str] | tuple[str, ...] | set[str],
    nome: str,
) -> list[str]:
    """
    Normaliza um texto unico ou uma colecao de textos.

    Args:
        valores (str | list | tuple | set):
            Valor ou colecao que sera normalizada.
        nome (str):
            Nome utilizado na mensagem de erro.

    Returns:
        list[str]:
            Textos preenchidos sem duplicidades.
    """
    # Converte uma string unica para lista
    if isinstance(valores, str):
        valores = [valores]

    # Valida as colecoes permitidas
    if not isinstance(valores, (list, tuple, set)) or not valores:
        raise TypeError(
            f"O parametro '{nome}' deve ser string ou colecao nao vazia."
        )

    # Valida e normaliza cada item
    resultado = []
    for valor in valores:
        texto = _validar_texto(valor, nome)
        if texto not in resultado:
            resultado.append(texto)

    return resultado


def _definir_propriedade(
    compromisso: Any,
    propriedade: str,
    valor: Any,
) -> None:
    """
    Define uma propriedade de um AppointmentItem com erro padronizado.

    Args:
        compromisso (Any):
            AppointmentItem que sera alterado.
        propriedade (str):
            Nome da propriedade COM.
        valor (Any):
            Valor aplicado a propriedade.

    Returns:
        None:
            A funcao apenas define a propriedade.
    """
    # Valida o compromisso antes da alteracao
    _validar_objeto(compromisso, "compromisso")

    try:
        # Define dinamicamente a propriedade COM
        setattr(compromisso, propriedade, valor)
    except Exception as error:
        raise CompromissoConfiguracaoError(
            f"Nao foi possivel definir a propriedade '{propriedade}'."
        ) from error


# ----------------------------------------------------------------------------
# OBTENCAO DO CALENDARIO E CRIACAO
# ----------------------------------------------------------------------------

def obter_pasta_calendario(namespace: Any) -> Any:
    """
    Retorna a pasta padrao de calendario do Namespace MAPI.

    Args:
        namespace (Any):
            Objeto NameSpace retornado pelo Outlook.

    Returns:
        Any:
            Objeto Folder correspondente ao calendario padrao.
    """
    # Valida o Namespace antes da chamada COM
    _validar_objeto(namespace, "namespace")

    try:
        # GetDefaultFolder recebe o valor olFolderCalendar
        return namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel obter a pasta padrao de calendario."
        ) from error


def criar_compromisso(outlook: Any) -> Any:
    """
    Cria um novo AppointmentItem no Outlook.

    Args:
        outlook (Any):
            Objeto Outlook.Application conectado.

    Returns:
        Any:
            Novo AppointmentItem ainda nao salvo.
    """
    # Valida a aplicacao recebida
    _validar_objeto(outlook, "outlook")

    try:
        # CreateItem recebe olAppointmentItem
        return outlook.CreateItem(OL_APPOINTMENT_ITEM)
    except Exception as error:
        raise CompromissoCriacaoError(
            "Nao foi possivel criar um compromisso no Outlook."
        ) from error


def configurar_compromisso(
    compromisso: Any,
    assunto: str,
    inicio: datetime | date,
    fim: datetime | date | None = None,
    duracao_minutos: int | None = None,
    local: str = "",
    corpo: str = "",
    dia_inteiro: bool = False,
    lembrete_minutos: int | None = 15,
    ocupacao: int = OL_BUSY,
    importancia: int = OL_IMPORTANCE_NORMAL,
    categorias: str | list[str] | tuple[str, ...] | set[str] | None = None,
) -> Any:
    """
    Configura as principais propriedades de um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem que sera configurado.
        assunto (str):
            Assunto exibido no calendario.
        inicio (datetime | date):
            Data e hora inicial.
        fim (datetime | date | None):
            Data e hora final opcional.
        duracao_minutos (int | None):
            Duracao usada quando fim nao e informado.
        local (str):
            Local do compromisso.
        corpo (str):
            Descricao textual.
        dia_inteiro (bool):
            Define se o evento ocupa o dia inteiro.
        lembrete_minutos (int | None):
            Antecedencia do lembrete ou None para desativar.
        ocupacao (int):
            Estado de disponibilidade.
        importancia (int):
            Nivel de importancia.
        categorias (str | list | tuple | set | None):
            Categorias associadas ao compromisso.

    Returns:
        Any:
            O mesmo AppointmentItem configurado.
    """
    # Define as propriedades textuais
    definir_assunto(compromisso, assunto)
    definir_corpo(compromisso, corpo)
    definir_local(compromisso, local)

    # Define datas por fim explicito ou duracao
    if fim is not None:
        definir_inicio_fim(compromisso, inicio, fim)
    else:
        definir_duracao(
            compromisso,
            inicio,
            60 if duracao_minutos is None else duracao_minutos,
        )

    # Define propriedades complementares
    definir_dia_inteiro(compromisso, dia_inteiro)
    definir_ocupacao(compromisso, ocupacao)
    definir_importancia(compromisso, importancia)

    # Ativa ou desativa o lembrete
    if lembrete_minutos is None:
        definir_lembrete(compromisso, ativar=False)
    else:
        definir_lembrete(
            compromisso,
            minutos_antes=lembrete_minutos,
            ativar=True,
        )

    # Define categorias quando informadas
    if categorias is not None:
        definir_categorias(compromisso, categorias)

    return compromisso


def criar_compromisso_completo(
    outlook: Any,
    assunto: str,
    inicio: datetime | date,
    fim: datetime | date | None = None,
    salvar: bool = True,
    exibir: bool = False,
    **configuracoes: Any,
) -> Any:
    """
    Cria, configura e opcionalmente salva ou exibe um compromisso.

    Args:
        outlook (Any):
            Objeto Outlook.Application conectado.
        assunto (str):
            Assunto do compromisso.
        inicio (datetime | date):
            Inicio do compromisso.
        fim (datetime | date | None):
            Fim opcional do compromisso.
        salvar (bool):
            Define se Save sera executado.
        exibir (bool):
            Define se Display sera executado.
        **configuracoes (Any):
            Propriedades encaminhadas para configurar_compromisso.

    Returns:
        Any:
            AppointmentItem criado.
    """
    # Valida as opcoes booleanas
    if not isinstance(salvar, bool) or not isinstance(exibir, bool):
        raise TypeError("Os parametros 'salvar' e 'exibir' devem ser booleanos.")

    # Cria e configura o item
    compromisso = criar_compromisso(outlook)
    configurar_compromisso(
        compromisso,
        assunto,
        inicio,
        fim=fim,
        **configuracoes,
    )

    # Executa as acoes finais solicitadas
    if salvar:
        salvar_compromisso(compromisso)
    if exibir:
        exibir_compromisso(compromisso)

    return compromisso


def criar_reuniao(
    outlook: Any,
    assunto: str,
    inicio: datetime | date,
    participantes_obrigatorios: str | list[str] | tuple[str, ...] | set[str],
    fim: datetime | date | None = None,
    participantes_opcionais: str | list[str] | tuple[str, ...] | set[str] | None = None,
    recursos: str | list[str] | tuple[str, ...] | set[str] | None = None,
    enviar: bool = False,
    exibir: bool = False,
    **configuracoes: Any,
) -> Any:
    """
    Cria uma reuniao com participantes obrigatorios e opcionais.

    Args:
        outlook (Any):
            Objeto Outlook.Application conectado.
        assunto (str):
            Assunto da reuniao.
        inicio (datetime | date):
            Inicio da reuniao.
        participantes_obrigatorios (str | list | tuple | set):
            Participantes obrigatorios.
        fim (datetime | date | None):
            Fim opcional da reuniao.
        participantes_opcionais (str | list | tuple | set | None):
            Participantes opcionais.
        recursos (str | list | tuple | set | None):
            Salas ou recursos.
        enviar (bool):
            Define se o convite sera enviado.
        exibir (bool):
            Define se a janela do convite sera exibida.
        **configuracoes (Any):
            Configuracoes adicionais do compromisso.

    Returns:
        Any:
            AppointmentItem configurado como reuniao.
    """
    # Valida as opcoes finais
    if not isinstance(enviar, bool) or not isinstance(exibir, bool):
        raise TypeError("Os parametros 'enviar' e 'exibir' devem ser booleanos.")

    # Cria o compromisso sem salva-lo antes da configuracao dos participantes
    reuniao = criar_compromisso(outlook)
    configurar_compromisso(
        reuniao,
        assunto,
        inicio,
        fim=fim,
        **configuracoes,
    )
    _definir_propriedade(reuniao, "MeetingStatus", OL_MEETING)

    # Adiciona os grupos de participantes
    adicionar_participantes(
        reuniao,
        participantes_obrigatorios,
        tipo=OL_REQUIRED,
    )
    if participantes_opcionais is not None:
        adicionar_participantes(
            reuniao,
            participantes_opcionais,
            tipo=OL_OPTIONAL,
        )
    if recursos is not None:
        adicionar_participantes(
            reuniao,
            recursos,
            tipo=OL_RESOURCE,
        )

    # Resolve todos os enderecos antes de enviar
    resolver_participantes(reuniao, exigir_resolucao=True)

    # Exibe, envia ou apenas salva a reuniao
    if exibir:
        exibir_compromisso(reuniao)
    if enviar:
        enviar_convite(reuniao)
    elif not exibir:
        salvar_compromisso(reuniao)

    return reuniao


# ----------------------------------------------------------------------------
# CONFIGURACAO DO COMPROMISSO
# ----------------------------------------------------------------------------

def definir_assunto(compromisso: Any, assunto: str) -> Any:
    """Define o assunto de um compromisso e retorna o proprio item."""
    _definir_propriedade(
        compromisso,
        "Subject",
        _validar_texto(assunto, "assunto"),
    )
    return compromisso


def definir_corpo(compromisso: Any, corpo: str) -> Any:
    """Define o corpo textual de um compromisso e retorna o item."""
    _definir_propriedade(
        compromisso,
        "Body",
        _validar_texto(corpo, "corpo", permitir_vazio=True),
    )
    return compromisso


def definir_local(compromisso: Any, local: str) -> Any:
    """Define o local de um compromisso e retorna o proprio item."""
    _definir_propriedade(
        compromisso,
        "Location",
        _validar_texto(local, "local", permitir_vazio=True),
    )
    return compromisso


def definir_inicio_fim(
    compromisso: Any,
    inicio: datetime | date,
    fim: datetime | date,
) -> Any:
    """
    Define o inicio e o fim de um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem que sera alterado.
        inicio (datetime | date):
            Data e hora inicial.
        fim (datetime | date):
            Data e hora final, posterior ao inicio.

    Returns:
        Any:
            O mesmo AppointmentItem.
    """
    # Converte e valida as datas
    data_inicio = _converter_datetime(inicio, "inicio")
    data_fim = _converter_datetime(fim, "fim")

    if data_fim <= data_inicio:
        raise ValueError("O fim deve ser posterior ao inicio.")

    # Aplica as datas ao item
    _definir_propriedade(compromisso, "Start", data_inicio)
    _definir_propriedade(compromisso, "End", data_fim)
    return compromisso


def definir_duracao(
    compromisso: Any,
    inicio: datetime | date,
    duracao_minutos: int,
) -> Any:
    """
    Define o inicio e a duracao em minutos de um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem que sera alterado.
        inicio (datetime | date):
            Data e hora inicial.
        duracao_minutos (int):
            Duracao positiva em minutos.

    Returns:
        Any:
            O mesmo AppointmentItem.
    """
    # Valida a duracao e converte o inicio
    _validar_inteiro(duracao_minutos, "duracao_minutos", minimo=1)
    data_inicio = _converter_datetime(inicio, "inicio")

    # Aplica inicio e duracao
    _definir_propriedade(compromisso, "Start", data_inicio)
    _definir_propriedade(compromisso, "Duration", duracao_minutos)
    return compromisso


def definir_dia_inteiro(compromisso: Any, dia_inteiro: bool = True) -> Any:
    """Define se o compromisso representa um evento de dia inteiro."""
    if not isinstance(dia_inteiro, bool):
        raise TypeError("O parametro 'dia_inteiro' deve ser booleano.")
    _definir_propriedade(compromisso, "AllDayEvent", dia_inteiro)
    return compromisso


def definir_lembrete(
    compromisso: Any,
    minutos_antes: int = 15,
    ativar: bool = True,
    som: bool = True,
) -> Any:
    """
    Configura o lembrete de um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem que sera alterado.
        minutos_antes (int):
            Antecedencia do lembrete em minutos.
        ativar (bool):
            Define se o lembrete esta ativo.
        som (bool):
            Define se o lembrete reproduzira som.

    Returns:
        Any:
            O mesmo AppointmentItem.
    """
    # Valida as opcoes
    if not isinstance(ativar, bool) or not isinstance(som, bool):
        raise TypeError("Os parametros 'ativar' e 'som' devem ser booleanos.")
    _validar_inteiro(minutos_antes, "minutos_antes", minimo=0)

    # Aplica as propriedades de lembrete
    _definir_propriedade(compromisso, "ReminderSet", ativar)
    _definir_propriedade(compromisso, "ReminderPlaySound", som and ativar)
    if ativar:
        _definir_propriedade(
            compromisso,
            "ReminderMinutesBeforeStart",
            minutos_antes,
        )
    return compromisso


def definir_ocupacao(compromisso: Any, ocupacao: int = OL_BUSY) -> Any:
    """Define o estado de disponibilidade exibido no calendario."""
    _validar_inteiro(ocupacao, "ocupacao", minimo=OL_FREE)
    permitidos = {
        OL_FREE,
        OL_TENTATIVE,
        OL_BUSY,
        OL_OUT_OF_OFFICE,
        OL_WORKING_ELSEWHERE,
    }
    if ocupacao not in permitidos:
        raise ValueError("O parametro 'ocupacao' possui valor invalido.")
    _definir_propriedade(compromisso, "BusyStatus", ocupacao)
    return compromisso


def definir_importancia(
    compromisso: Any,
    importancia: int = OL_IMPORTANCE_NORMAL,
) -> Any:
    """Define o nivel de importancia do compromisso."""
    _validar_inteiro(importancia, "importancia", minimo=0)
    if importancia not in {
        OL_IMPORTANCE_LOW,
        OL_IMPORTANCE_NORMAL,
        OL_IMPORTANCE_HIGH,
    }:
        raise ValueError("O parametro 'importancia' possui valor invalido.")
    _definir_propriedade(compromisso, "Importance", importancia)
    return compromisso


def definir_categorias(
    compromisso: Any,
    categorias: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define uma ou mais categorias separadas pelo padrao do Outlook."""
    valores = _normalizar_lista_textos(categorias, "categorias")
    _definir_propriedade(compromisso, "Categories", ", ".join(valores))
    return compromisso


# ----------------------------------------------------------------------------
# PARTICIPANTES
# ----------------------------------------------------------------------------

def adicionar_participante(
    compromisso: Any,
    participante: str,
    tipo: int = OL_REQUIRED,
) -> Any:
    """
    Adiciona um participante obrigatorio, opcional ou recurso.

    Args:
        compromisso (Any):
            AppointmentItem que recebera o participante.
        participante (str):
            Nome, alias ou endereco SMTP.
        tipo (int):
            OL_REQUIRED, OL_OPTIONAL ou OL_RESOURCE.

    Returns:
        Any:
            Objeto Recipient criado.
    """
    # Valida o compromisso, endereco e tipo
    _validar_objeto(compromisso, "compromisso")
    endereco = _validar_texto(participante, "participante")
    _validar_inteiro(tipo, "tipo", minimo=OL_ORGANIZER)

    if tipo not in {OL_REQUIRED, OL_OPTIONAL, OL_RESOURCE}:
        raise ValueError("O tipo do participante nao e permitido.")

    try:
        # Adiciona o destinatario e define seu tipo
        recipient = compromisso.Recipients.Add(endereco)
        recipient.Type = tipo
        return recipient
    except Exception as error:
        raise ParticipanteInvalidoError(
            f"Nao foi possivel adicionar o participante '{endereco}'."
        ) from error


def adicionar_participantes(
    compromisso: Any,
    participantes: str | list[str] | tuple[str, ...] | set[str],
    tipo: int = OL_REQUIRED,
) -> list[Any]:
    """
    Adiciona varios participantes ao compromisso.

    Args:
        compromisso (Any):
            AppointmentItem que recebera os participantes.
        participantes (str | list | tuple | set):
            Enderecos ou nomes dos participantes.
        tipo (int):
            Tipo aplicado a todos os participantes.

    Returns:
        list[Any]:
            Objetos Recipient adicionados.
    """
    # Normaliza a colecao para evitar duplicidades
    valores = _normalizar_lista_textos(participantes, "participantes")

    # Adiciona cada participante individualmente
    return [
        adicionar_participante(compromisso, valor, tipo)
        for valor in valores
    ]


def resolver_participantes(
    compromisso: Any,
    exigir_resolucao: bool = True,
) -> bool:
    """
    Solicita ao Outlook a resolucao dos participantes da reuniao.

    Args:
        compromisso (Any):
            AppointmentItem cujos participantes serao resolvidos.
        exigir_resolucao (bool):
            Gera erro quando algum participante nao pode ser resolvido.

    Returns:
        bool:
            Resultado retornado por ResolveAll.
    """
    # Valida o item e a opcao
    _validar_objeto(compromisso, "compromisso")
    if not isinstance(exigir_resolucao, bool):
        raise TypeError("O parametro 'exigir_resolucao' deve ser booleano.")

    try:
        # ResolveAll retorna True quando todos os nomes foram resolvidos
        resolvido = bool(compromisso.Recipients.ResolveAll())
    except Exception as error:
        raise ParticipanteInvalidoError(
            "Nao foi possivel resolver os participantes da reuniao."
        ) from error

    # Gera erro quando a resolucao completa e obrigatoria
    if exigir_resolucao and not resolvido:
        raise ParticipanteInvalidoError(
            "Um ou mais participantes nao foram resolvidos pelo Outlook."
        )

    return resolvido


def listar_participantes(compromisso: Any) -> list[dict[str, Any]]:
    """
    Retorna os metadados dos participantes de uma reuniao.

    Args:
        compromisso (Any):
            AppointmentItem que sera consultado.

    Returns:
        list[dict[str, Any]]:
            Indice, nome, endereco, tipo e estado de resolucao.
    """
    # Valida o item e recupera a colecao
    _validar_objeto(compromisso, "compromisso")

    try:
        recipients = compromisso.Recipients
        quantidade = int(recipients.Count)
    except Exception as error:
        raise ParticipanteInvalidoError(
            "Nao foi possivel consultar os participantes."
        ) from error

    # Percorre a colecao COM iniciada em um
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            recipient = recipients.Item(indice)
            address_entry = _obter_propriedade(recipient, "AddressEntry")
            endereco = _obter_propriedade(address_entry, "Address", "")
            resultado.append(
                {
                    "indice": indice,
                    "nome": str(_obter_propriedade(recipient, "Name", "")),
                    "endereco": str(endereco or ""),
                    "tipo": _obter_propriedade(recipient, "Type"),
                    "resolvido": bool(
                        _obter_propriedade(recipient, "Resolved", False)
                    ),
                }
            )
        except Exception as error:
            raise ParticipanteInvalidoError(
                f"Nao foi possivel consultar o participante {indice}."
            ) from error

    return resultado


def remover_participante(compromisso: Any, indice: int) -> bool:
    """Remove um participante pelo indice COM iniciado em um."""
    _validar_objeto(compromisso, "compromisso")
    _validar_inteiro(indice, "indice", minimo=1)

    try:
        compromisso.Recipients.Remove(indice)
        return True
    except Exception as error:
        raise ParticipanteInvalidoError(
            f"Nao foi possivel remover o participante {indice}."
        ) from error


def remover_todos_participantes(compromisso: Any) -> int:
    """Remove todos os participantes e retorna a quantidade removida."""
    _validar_objeto(compromisso, "compromisso")

    try:
        quantidade = int(compromisso.Recipients.Count)
    except Exception as error:
        raise ParticipanteInvalidoError(
            "Nao foi possivel contar os participantes."
        ) from error

    # Remove do ultimo para o primeiro para evitar deslocamento de indices
    for indice in range(quantidade, 0, -1):
        remover_participante(compromisso, indice)

    return quantidade


# ----------------------------------------------------------------------------
# ACOES DO COMPROMISSO
# ----------------------------------------------------------------------------

def salvar_compromisso(compromisso: Any) -> Any:
    """Salva o compromisso e retorna o proprio AppointmentItem."""
    _validar_objeto(compromisso, "compromisso")
    try:
        compromisso.Save()
        return compromisso
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel salvar o compromisso."
        ) from error


def exibir_compromisso(compromisso: Any, modal: bool = False) -> Any:
    """Exibe a janela do compromisso no Outlook."""
    _validar_objeto(compromisso, "compromisso")
    if not isinstance(modal, bool):
        raise TypeError("O parametro 'modal' deve ser booleano.")
    try:
        compromisso.Display(modal)
        return compromisso
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel exibir o compromisso."
        ) from error


def enviar_convite(
    compromisso: Any,
    conta_envio: Any | None = None,
    resolver: bool = True,
) -> Any:
    """
    Envia uma solicitacao de reuniao.

    Args:
        compromisso (Any):
            AppointmentItem configurado como reuniao.
        conta_envio (Any | None):
            Conta opcional aplicada em SendUsingAccount.
        resolver (bool):
            Define se os participantes serao resolvidos antes do envio.

    Returns:
        Any:
            O mesmo AppointmentItem enviado.
    """
    # Valida o item e a opcao de resolucao
    _validar_objeto(compromisso, "compromisso")
    if not isinstance(resolver, bool):
        raise TypeError("O parametro 'resolver' deve ser booleano.")

    # Garante que o item representa uma reuniao
    _definir_propriedade(compromisso, "MeetingStatus", OL_MEETING)

    # Resolve destinatarios antes do envio quando solicitado
    if resolver:
        resolver_participantes(compromisso, exigir_resolucao=True)

    # Define uma conta especifica quando fornecida
    if conta_envio is not None:
        _definir_propriedade(compromisso, "SendUsingAccount", conta_envio)

    try:
        compromisso.Send()
        return compromisso
    except Exception as error:
        raise CompromissoEnvioError(
            "Nao foi possivel enviar o convite de reuniao."
        ) from error


def cancelar_compromisso(
    compromisso: Any,
    enviar_cancelamento: bool = True,
) -> Any:
    """
    Marca a reuniao como cancelada e opcionalmente envia o cancelamento.

    Args:
        compromisso (Any):
            AppointmentItem que sera cancelado.
        enviar_cancelamento (bool):
            Define se Send sera executado apos o cancelamento.

    Returns:
        Any:
            O mesmo AppointmentItem cancelado.
    """
    # Valida a opcao de envio
    if not isinstance(enviar_cancelamento, bool):
        raise TypeError(
            "O parametro 'enviar_cancelamento' deve ser booleano."
        )

    # Altera o estado para reuniao cancelada
    _definir_propriedade(
        compromisso,
        "MeetingStatus",
        OL_MEETING_CANCELED,
    )

    try:
        # Envia a notificacao ou apenas salva a alteracao
        if enviar_cancelamento:
            compromisso.Send()
        else:
            compromisso.Save()
        return compromisso
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel cancelar o compromisso."
        ) from error


def excluir_compromisso(compromisso: Any) -> bool:
    """Exclui um compromisso da pasta onde ele esta armazenado."""
    _validar_objeto(compromisso, "compromisso")
    try:
        compromisso.Delete()
        return True
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel excluir o compromisso."
        ) from error


# ----------------------------------------------------------------------------
# CONSULTA E PESQUISA
# ----------------------------------------------------------------------------

def listar_compromissos(
    pasta_calendario: Any,
    incluir_recorrencias: bool = True,
    ordenar_por_inicio: bool = True,
    ordem_decrescente: bool = False,
    limite: int | None = None,
) -> list[Any]:
    """
    Lista os compromissos existentes em uma pasta de calendario.

    Args:
        pasta_calendario (Any):
            Pasta Calendar que sera consultada.
        incluir_recorrencias (bool):
            Define IncludeRecurrences na colecao Items.
        ordenar_por_inicio (bool):
            Define se os itens serao ordenados por Start.
        ordem_decrescente (bool):
            Define a direcao da ordenacao.
        limite (int | None):
            Quantidade maxima de itens retornados.

    Returns:
        list[Any]:
            Itens de calendario encontrados.
    """
    # Valida a pasta e as opcoes
    _validar_objeto(pasta_calendario, "pasta_calendario")
    for nome, valor in {
        "incluir_recorrencias": incluir_recorrencias,
        "ordenar_por_inicio": ordenar_por_inicio,
        "ordem_decrescente": ordem_decrescente,
    }.items():
        if not isinstance(valor, bool):
            raise TypeError(f"O parametro '{nome}' deve ser booleano.")

    if limite is not None:
        _validar_inteiro(limite, "limite", minimo=1)

    try:
        # Recupera e configura a colecao Items
        itens = pasta_calendario.Items
        itens.IncludeRecurrences = incluir_recorrencias
        if ordenar_por_inicio:
            itens.Sort("[Start]", ordem_decrescente)
        quantidade = int(itens.Count)
    except Exception as error:
        raise CalendarioOutlookError(
            "Nao foi possivel listar os compromissos do calendario."
        ) from error

    # Converte a colecao COM baseada em um para uma lista Python
    quantidade_retorno = quantidade if limite is None else min(quantidade, limite)
    resultado = []
    for indice in range(1, quantidade_retorno + 1):
        try:
            resultado.append(itens.Item(indice))
        except Exception as error:
            raise CalendarioOutlookError(
                f"Nao foi possivel obter o compromisso {indice}."
            ) from error

    return resultado


def buscar_compromissos(
    pasta_calendario: Any,
    inicio: datetime | date,
    fim: datetime | date,
    assunto_contem: str | None = None,
    local_contem: str | None = None,
    limite: int | None = None,
) -> list[Any]:
    """
    Busca compromissos dentro de um intervalo e aplica filtros textuais.

    Args:
        pasta_calendario (Any):
            Pasta Calendar que sera consultada.
        inicio (datetime | date):
            Inicio do intervalo.
        fim (datetime | date):
            Fim do intervalo.
        assunto_contem (str | None):
            Trecho opcional procurado no assunto.
        local_contem (str | None):
            Trecho opcional procurado no local.
        limite (int | None):
            Quantidade maxima de itens retornados.

    Returns:
        list[Any]:
            Compromissos correspondentes.
    """
    # Valida as datas do intervalo
    data_inicio = _converter_datetime(inicio, "inicio")
    data_fim = _converter_datetime(fim, "fim")
    if data_fim <= data_inicio:
        raise ValueError("O fim do intervalo deve ser posterior ao inicio.")

    # Valida os filtros opcionais
    assunto = None
    local = None
    if assunto_contem is not None:
        assunto = _validar_texto(assunto_contem, "assunto_contem").lower()
    if local_contem is not None:
        local = _validar_texto(local_contem, "local_contem").lower()
    if limite is not None:
        _validar_inteiro(limite, "limite", minimo=1)

    # Recupera todos os itens de forma ordenada
    itens = listar_compromissos(
        pasta_calendario,
        incluir_recorrencias=True,
        ordenar_por_inicio=True,
    )

    # Filtra sobreposicao de datas e propriedades textuais
    resultado = []
    for item in itens:
        item_inicio = _obter_propriedade(item, "Start")
        item_fim = _obter_propriedade(item, "End")

        if not isinstance(item_inicio, (datetime, date)):
            continue
        item_inicio = _converter_datetime(item_inicio, "Start")
        item_fim = (
            _converter_datetime(item_fim, "End")
            if isinstance(item_fim, (datetime, date))
            else item_inicio
        )

        # Inclui itens que sobrepoem o intervalo solicitado
        if item_fim < data_inicio or item_inicio > data_fim:
            continue

        item_assunto = str(_obter_propriedade(item, "Subject", "")).lower()
        item_local = str(_obter_propriedade(item, "Location", "")).lower()

        if assunto is not None and assunto not in item_assunto:
            continue
        if local is not None and local not in item_local:
            continue

        resultado.append(item)
        if limite is not None and len(resultado) >= limite:
            break

    return resultado


def obter_compromisso_por_entry_id(
    namespace: Any,
    entry_id: str,
    store_id: str | None = None,
) -> Any:
    """
    Localiza um compromisso por EntryID e StoreID opcional.

    Args:
        namespace (Any):
            Objeto NameSpace MAPI.
        entry_id (str):
            Identificador persistente do item.
        store_id (str | None):
            Identificador opcional da caixa postal.

    Returns:
        Any:
            Item localizado pelo Outlook.
    """
    # Valida os parametros obrigatorios
    _validar_objeto(namespace, "namespace")
    identificador = _validar_texto(entry_id, "entry_id")

    if store_id is not None:
        store_id = _validar_texto(store_id, "store_id")

    try:
        # GetItemFromID aceita StoreID como segundo argumento opcional
        item = (
            namespace.GetItemFromID(identificador, store_id)
            if store_id is not None
            else namespace.GetItemFromID(identificador)
        )
    except Exception as error:
        raise CompromissoNaoEncontradoError(
            f"Nao foi possivel localizar o compromisso '{identificador}'."
        ) from error

    if item is None:
        raise CompromissoNaoEncontradoError(
            f"O compromisso '{identificador}' nao foi encontrado."
        )

    return item


def obter_dados_compromisso(compromisso: Any) -> dict[str, Any]:
    """
    Converte as principais propriedades do compromisso para dicionario.

    Args:
        compromisso (Any):
            AppointmentItem que sera convertido.

    Returns:
        dict[str, Any]:
            Dados estruturados do compromisso.
    """
    # Valida o item antes da leitura
    _validar_objeto(compromisso, "compromisso")

    # Retorna propriedades comuns de AppointmentItem
    return {
        "entry_id": _obter_propriedade(compromisso, "EntryID"),
        "assunto": _obter_propriedade(compromisso, "Subject", ""),
        "corpo": _obter_propriedade(compromisso, "Body", ""),
        "local": _obter_propriedade(compromisso, "Location", ""),
        "inicio": _obter_propriedade(compromisso, "Start"),
        "fim": _obter_propriedade(compromisso, "End"),
        "duracao": _obter_propriedade(compromisso, "Duration"),
        "dia_inteiro": bool(
            _obter_propriedade(compromisso, "AllDayEvent", False)
        ),
        "lembrete_ativo": bool(
            _obter_propriedade(compromisso, "ReminderSet", False)
        ),
        "lembrete_minutos": _obter_propriedade(
            compromisso,
            "ReminderMinutesBeforeStart",
        ),
        "ocupacao": _obter_propriedade(compromisso, "BusyStatus"),
        "importancia": _obter_propriedade(compromisso, "Importance"),
        "categorias": _obter_propriedade(compromisso, "Categories", ""),
        "organizador": _obter_propriedade(compromisso, "Organizer", ""),
        "status_reuniao": _obter_propriedade(
            compromisso,
            "MeetingStatus",
        ),
        "recorrente": bool(
            _obter_propriedade(compromisso, "IsRecurring", False)
        ),
    }


def compromisso_possui_conflito(compromisso: Any) -> bool:
    """Verifica se o Outlook marcou o compromisso como conflito."""
    _validar_objeto(compromisso, "compromisso")
    return bool(_obter_propriedade(compromisso, "IsConflict", False))


# ----------------------------------------------------------------------------
# RECORRENCIA
# ----------------------------------------------------------------------------

def definir_recorrencia_diaria(
    compromisso: Any,
    intervalo_dias: int = 1,
    data_fim: datetime | date | None = None,
    quantidade_ocorrencias: int | None = None,
) -> Any:
    """
    Configura recorrencia diaria para um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem recorrente.
        intervalo_dias (int):
            Intervalo entre ocorrencias.
        data_fim (datetime | date | None):
            Data final opcional.
        quantidade_ocorrencias (int | None):
            Quantidade opcional de ocorrencias.

    Returns:
        Any:
            Objeto RecurrencePattern configurado.
    """
    # Valida os parametros antes de obter o padrao
    _validar_objeto(compromisso, "compromisso")
    _validar_inteiro(intervalo_dias, "intervalo_dias", minimo=1)

    if quantidade_ocorrencias is not None:
        _validar_inteiro(
            quantidade_ocorrencias,
            "quantidade_ocorrencias",
            minimo=1,
        )

    try:
        padrao = compromisso.GetRecurrencePattern()
        padrao.RecurrenceType = OL_RECURRENCE_DAILY
        padrao.Interval = intervalo_dias

        if data_fim is not None:
            padrao.PatternEndDate = _converter_datetime(
                data_fim,
                "data_fim",
            )
        elif quantidade_ocorrencias is not None:
            padrao.Occurrences = quantidade_ocorrencias

        return padrao
    except Exception as error:
        raise RecorrenciaError(
            "Nao foi possivel configurar a recorrencia diaria."
        ) from error


def definir_recorrencia_semanal(
    compromisso: Any,
    mascara_dias_semana: int,
    intervalo_semanas: int = 1,
    data_fim: datetime | date | None = None,
    quantidade_ocorrencias: int | None = None,
) -> Any:
    """
    Configura recorrencia semanal para um compromisso.

    Args:
        compromisso (Any):
            AppointmentItem recorrente.
        mascara_dias_semana (int):
            Mascara OlDaysOfWeek dos dias da semana.
        intervalo_semanas (int):
            Intervalo entre as semanas.
        data_fim (datetime | date | None):
            Data final opcional.
        quantidade_ocorrencias (int | None):
            Quantidade opcional de ocorrencias.

    Returns:
        Any:
            Objeto RecurrencePattern configurado.
    """
    # Valida valores inteiros positivos
    _validar_objeto(compromisso, "compromisso")
    _validar_inteiro(mascara_dias_semana, "mascara_dias_semana", minimo=1)
    _validar_inteiro(intervalo_semanas, "intervalo_semanas", minimo=1)

    if quantidade_ocorrencias is not None:
        _validar_inteiro(
            quantidade_ocorrencias,
            "quantidade_ocorrencias",
            minimo=1,
        )

    try:
        padrao = compromisso.GetRecurrencePattern()
        padrao.RecurrenceType = OL_RECURRENCE_WEEKLY
        padrao.DayOfWeekMask = mascara_dias_semana
        padrao.Interval = intervalo_semanas

        if data_fim is not None:
            padrao.PatternEndDate = _converter_datetime(data_fim, "data_fim")
        elif quantidade_ocorrencias is not None:
            padrao.Occurrences = quantidade_ocorrencias

        return padrao
    except Exception as error:
        raise RecorrenciaError(
            "Nao foi possivel configurar a recorrencia semanal."
        ) from error


def definir_recorrencia_mensal(
    compromisso: Any,
    dia_mes: int,
    intervalo_meses: int = 1,
    data_fim: datetime | date | None = None,
    quantidade_ocorrencias: int | None = None,
) -> Any:
    """
    Configura recorrencia mensal em um dia fixo do mes.

    Args:
        compromisso (Any):
            AppointmentItem recorrente.
        dia_mes (int):
            Dia do mes entre 1 e 31.
        intervalo_meses (int):
            Intervalo entre os meses.
        data_fim (datetime | date | None):
            Data final opcional.
        quantidade_ocorrencias (int | None):
            Quantidade opcional de ocorrencias.

    Returns:
        Any:
            Objeto RecurrencePattern configurado.
    """
    # Valida o dia e o intervalo
    _validar_objeto(compromisso, "compromisso")
    _validar_inteiro(dia_mes, "dia_mes", minimo=1)
    _validar_inteiro(intervalo_meses, "intervalo_meses", minimo=1)

    if dia_mes > 31:
        raise ValueError("O parametro 'dia_mes' deve estar entre 1 e 31.")

    if quantidade_ocorrencias is not None:
        _validar_inteiro(
            quantidade_ocorrencias,
            "quantidade_ocorrencias",
            minimo=1,
        )

    try:
        padrao = compromisso.GetRecurrencePattern()
        padrao.RecurrenceType = OL_RECURRENCE_MONTHLY
        padrao.DayOfMonth = dia_mes
        padrao.Interval = intervalo_meses

        if data_fim is not None:
            padrao.PatternEndDate = _converter_datetime(data_fim, "data_fim")
        elif quantidade_ocorrencias is not None:
            padrao.Occurrences = quantidade_ocorrencias

        return padrao
    except Exception as error:
        raise RecorrenciaError(
            "Nao foi possivel configurar a recorrencia mensal."
        ) from error


def remover_recorrencia(compromisso: Any) -> Any:
    """Remove o padrao de recorrencia e retorna o compromisso."""
    _validar_objeto(compromisso, "compromisso")
    try:
        compromisso.ClearRecurrencePattern()
        return compromisso
    except Exception as error:
        raise RecorrenciaError(
            "Nao foi possivel remover a recorrencia do compromisso."
        ) from error
