"""
============================================================
Modulo: EXCEL / planilhas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 29/09/2026
Ultima Alteracao: 29/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para criar, consultar, renomear,
    copiar, mover, ocultar, exibir, limpar e remover planilhas em arquivos
    Excel utilizando openpyxl. O modulo tambem permite ordenar abas,
    definir a planilha ativa, controlar cores das guias, proteger abas e
    consultar informacoes estruturais.
Funcoes disponiveis:
    - listar_planilhas()
    - contar_planilhas()
    - planilha_existe()
    - obter_planilha()
    - obter_planilha_ativa()
    - definir_planilha_ativa()
    - criar_planilha()
    - renomear_planilha()
    - copiar_planilha()
    - copiar_planilha_entre_workbooks()
    - mover_planilha()
    - ordenar_planilhas()
    - remover_planilha()
    - remover_planilhas()
    - limpar_planilha()
    - ocultar_planilha()
    - exibir_planilha()
    - definir_cor_guia()
    - proteger_planilha()
    - desproteger_planilha()
    - obter_dimensoes_planilha()
    - obter_informacoes_planilha()
Dependencias:
    - openpyxl
Historico:
    v1.0.0 - 29/09/2026
        - Criacao inicial do modulo.
        - Inclusao das operacoes de consulta e criacao de planilhas.
        - Inclusao das operacoes de copia, movimentacao e ordenacao.
        - Inclusao das operacoes de ocultacao, protecao e remocao.
        - Inclusao das consultas de dimensoes e informacoes estruturais.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa copy para duplicar estilos e objetos sem referencias compartilhadas
from copy import copy

# Importa Iterable para validar colecoes de nomes de planilhas
from collections.abc import Iterable

# Importa Any para anotacoes de valores flexiveis
from typing import Any

# Importa Workbook para validar e manipular arquivos Excel em memoria
from openpyxl import Workbook

# Importa MergedCell para ignorar celulas mescladas secundarias
from openpyxl.cell.cell import MergedCell

# Importa get_column_letter para consultar a ultima coluna utilizada
from openpyxl.utils import get_column_letter

# Importa StyleArray para restaurar o estilo padrao das celulas
from openpyxl.styles.cell_style import StyleArray

# Importa Worksheet para validar e manipular planilhas
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define o nome padrao utilizado na criacao de uma planilha
NOME_PLANILHA_PADRAO = "Planilha"

# Define o limite de caracteres permitido pelo Excel no nome da planilha
LIMITE_CARACTERES_PLANILHA = 31

# Define os caracteres proibidos nos nomes das planilhas
CARACTERES_INVALIDOS_PLANILHA = {
    "[",
    "]",
    ":",
    "*",
    "?",
    "/",
    "\\"
}

# Define os estados de visibilidade reconhecidos pelo openpyxl
ESTADOS_VISIBILIDADE = {
    "visible",
    "hidden",
    "veryHidden"
}

# Define a cor padrao utilizada nas guias quando solicitada
COR_GUIA_PADRAO = "FF4472C4"


# ----------------------------------------------------------------------------
# EXCECOES
# ----------------------------------------------------------------------------

class ExcelPlanilhasError(Exception):
    """Erro base para operacoes relacionadas a planilhas Excel."""


class WorkbookPlanilhasInvalidoError(ExcelPlanilhasError, TypeError):
    """Indica que o objeto informado nao e um Workbook valido."""


class PlanilhaInvalidaError(ExcelPlanilhasError, TypeError):
    """Indica que o objeto informado nao e uma Worksheet valida."""


class NomePlanilhaInvalidoError(ExcelPlanilhasError, ValueError):
    """Indica que o nome informado nao atende as regras do Excel."""


class PlanilhaNaoEncontradaError(ExcelPlanilhasError, KeyError):
    """Indica que a planilha solicitada nao existe no Workbook."""


class PlanilhaExistenteError(ExcelPlanilhasError, FileExistsError):
    """Indica que uma planilha com o mesmo nome ja existe no Workbook."""


class OperacaoPlanilhaInvalidaError(ExcelPlanilhasError, ValueError):
    """Indica que uma operacao nao pode ser aplicada a planilha informada."""


class CorGuiaInvalidaError(ExcelPlanilhasError, ValueError):
    """Indica que a cor informada para a guia da planilha e invalida."""


# ----------------------------------------------------------------------------
# FUNCOES INTERNAS DE VALIDACAO
# ----------------------------------------------------------------------------

def _validar_booleano(
    valor: bool,
    nome: str
) -> bool:
    """
    Valida estritamente um parametro booleano.

    Args:
        valor (bool):
            Valor que sera validado.
        nome (str):
            Nome do parametro utilizado na mensagem de erro.

    Returns:
        bool:
            O proprio valor booleano validado.

    Raises:
        TypeError:
            Caso o valor nao seja booleano.
    """
    # Verifica se o valor informado e booleano
    if not isinstance(valor, bool):
        raise TypeError(
            f"O parametro '{nome}' deve ser booleano."
        )

    # Retorna o valor validado
    return valor


def _validar_inteiro(
    valor: int,
    nome: str,
    minimo: int | None = None,
    maximo: int | None = None
) -> int:
    """
    Valida um numero inteiro e seus limites opcionais.

    Args:
        valor (int):
            Numero inteiro que sera validado.
        nome (str):
            Nome do parametro utilizado na mensagem de erro.
        minimo (int | None):
            Menor valor permitido.
        maximo (int | None):
            Maior valor permitido.

    Returns:
        int:
            O proprio numero inteiro validado.

    Raises:
        TypeError:
            Caso o valor nao seja inteiro.
        ValueError:
            Caso o valor esteja fora dos limites permitidos.
    """
    # Impede que booleanos sejam interpretados como inteiros
    if isinstance(valor, bool) or not isinstance(valor, int):
        raise TypeError(
            f"O parametro '{nome}' deve ser inteiro."
        )

    # Verifica o limite minimo quando informado
    if minimo is not None and valor < minimo:
        raise ValueError(
            f"O parametro '{nome}' deve ser maior ou igual a {minimo}."
        )

    # Verifica o limite maximo quando informado
    if maximo is not None and valor > maximo:
        raise ValueError(
            f"O parametro '{nome}' deve ser menor ou igual a {maximo}."
        )

    # Retorna o inteiro validado
    return valor


def _validar_texto(
    valor: str,
    nome: str
) -> str:
    """
    Valida e normaliza um parametro textual obrigatorio.

    Args:
        valor (str):
            Texto que sera validado.
        nome (str):
            Nome do parametro utilizado na mensagem de erro.

    Returns:
        str:
            Texto sem espacos externos.

    Raises:
        TypeError:
            Caso o valor nao seja uma string.
        ValueError:
            Caso o texto esteja vazio.
    """
    # Verifica se o valor informado e uma string
    if not isinstance(valor, str):
        raise TypeError(
            f"O parametro '{nome}' deve ser uma string."
        )

    # Remove os espacos externos
    texto = valor.strip()

    # Rejeita textos vazios
    if not texto:
        raise ValueError(
            f"O parametro '{nome}' nao pode estar vazio."
        )

    # Retorna o texto validado
    return texto


def _validar_workbook(
    workbook: Workbook
) -> Workbook:
    """
    Valida se o objeto informado e um Workbook do openpyxl.

    Args:
        workbook (Workbook):
            Workbook que sera validado.

    Returns:
        Workbook:
            O proprio Workbook validado.

    Raises:
        WorkbookPlanilhasInvalidoError:
            Caso o objeto nao seja um Workbook.
    """
    # Verifica o tipo do objeto recebido
    if not isinstance(workbook, Workbook):
        raise WorkbookPlanilhasInvalidoError(
            "O objeto fornecido nao e um Workbook do openpyxl."
        )

    # Retorna o Workbook validado
    return workbook


def _validar_planilha(
    planilha: Worksheet
) -> Worksheet:
    """
    Valida se o objeto informado e uma Worksheet do openpyxl.

    Args:
        planilha (Worksheet):
            Planilha que sera validada.

    Returns:
        Worksheet:
            A propria planilha validada.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Verifica o tipo do objeto recebido
    if not isinstance(planilha, Worksheet):
        raise PlanilhaInvalidaError(
            "O objeto fornecido nao e uma planilha do openpyxl."
        )

    # Retorna a planilha validada
    return planilha


def _validar_nome_planilha(
    nome: str
) -> str:
    """
    Valida um nome conforme as regras de planilhas do Excel.

    Args:
        nome (str):
            Nome da planilha que sera validado.

    Returns:
        str:
            Nome validado sem espacos externos.

    Raises:
        NomePlanilhaInvalidoError:
            Caso o nome seja longo, vazio ou possua caracteres proibidos.
    """
    # Valida e normaliza o texto do nome
    try:
        nome_validado = _validar_texto(
            nome,
            "nome"
        )
    except (TypeError, ValueError) as erro:
        raise NomePlanilhaInvalidoError(
            "O nome da planilha e invalido."
        ) from erro

    # Verifica o limite de caracteres do Excel
    if len(nome_validado) > LIMITE_CARACTERES_PLANILHA:
        raise NomePlanilhaInvalidoError(
            "O nome da planilha nao pode ultrapassar 31 caracteres."
        )

    # Localiza caracteres proibidos no nome
    encontrados = sorted(
        caractere
        for caractere in CARACTERES_INVALIDOS_PLANILHA
        if caractere in nome_validado
    )

    # Rejeita caracteres proibidos
    if encontrados:
        raise NomePlanilhaInvalidoError(
            f"O nome da planilha possui caracteres invalidos: {encontrados}."
        )

    # Rejeita nomes iniciados ou finalizados por apostrofo
    if nome_validado.startswith("'") or nome_validado.endswith("'"):
        raise NomePlanilhaInvalidoError(
            "O nome da planilha nao pode comecar ou terminar com apostrofo."
        )

    # Retorna o nome validado
    return nome_validado


def _normalizar_cor(
    cor: str
) -> str:
    """
    Valida e normaliza uma cor hexadecimal para o formato ARGB.

    Args:
        cor (str):
            Cor hexadecimal RGB ou ARGB.

    Returns:
        str:
            Codigo ARGB com oito caracteres.

    Raises:
        CorGuiaInvalidaError:
            Caso a cor possua tamanho ou caracteres invalidos.
    """
    # Valida o texto e remove o prefixo opcional
    try:
        codigo = _validar_texto(
            cor,
            "cor"
        ).lstrip("#").upper()
    except (TypeError, ValueError) as erro:
        raise CorGuiaInvalidaError(
            "O codigo de cor informado e invalido."
        ) from erro

    # Verifica o tamanho do codigo
    if len(codigo) not in {6, 8}:
        raise CorGuiaInvalidaError(
            "A cor deve possuir seis caracteres RGB ou oito caracteres ARGB."
        )

    # Verifica se todos os caracteres sao hexadecimais
    if any(
        caractere not in "0123456789ABCDEF"
        for caractere in codigo
    ):
        raise CorGuiaInvalidaError(
            f"O codigo de cor e invalido: '{cor}'."
        )

    # Adiciona o canal alfa nas cores RGB
    if len(codigo) == 6:
        codigo = f"FF{codigo}"

    # Retorna a cor normalizada
    return codigo


def _obter_planilha_por_referencia(
    workbook: Workbook,
    referencia: str | int | Worksheet
) -> Worksheet:
    """
    Obtem uma planilha por nome, indice ou objeto Worksheet.

    Args:
        workbook (Workbook):
            Workbook que contem a planilha.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha.

    Returns:
        Worksheet:
            Planilha localizada no Workbook.

    Raises:
        PlanilhaNaoEncontradaError:
            Caso a referencia nao corresponda a uma planilha do Workbook.
    """
    # Valida o Workbook recebido
    _validar_workbook(workbook)

    # Trata uma referencia por objeto Worksheet
    if isinstance(referencia, Worksheet):
        if referencia not in workbook.worksheets:
            raise PlanilhaNaoEncontradaError(
                "A planilha informada nao pertence ao Workbook."
            )
        return referencia

    # Trata uma referencia textual
    if isinstance(referencia, str):
        nome = _validar_nome_planilha(referencia)

        if nome not in workbook.sheetnames:
            raise PlanilhaNaoEncontradaError(
                f"A planilha '{nome}' nao foi encontrada. "
                f"Planilhas disponiveis: {workbook.sheetnames}."
            )

        return workbook[nome]

    # Trata uma referencia numerica
    if isinstance(referencia, int) and not isinstance(referencia, bool):
        indice = _validar_inteiro(
            referencia,
            "referencia",
            minimo=0
        )

        if indice >= len(workbook.worksheets):
            raise PlanilhaNaoEncontradaError(
                f"O indice da planilha e invalido: {indice}."
            )

        return workbook.worksheets[indice]

    # Rejeita tipos incompativeis
    raise TypeError(
        "A referencia da planilha deve ser string, inteiro ou Worksheet."
    )


# ----------------------------------------------------------------------------
# FUNCOES DE CONSULTA
# ----------------------------------------------------------------------------

def listar_planilhas(
    workbook: Workbook,
    incluir_ocultas: bool = True
) -> list[str]:
    """
    Lista os nomes das planilhas existentes no Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera consultado.
        incluir_ocultas (bool):
            Define se planilhas ocultas serao incluidas.

    Returns:
        list[str]:
            Nomes das planilhas na ordem do Workbook.

    Raises:
        WorkbookPlanilhasInvalidoError:
            Caso o objeto nao seja um Workbook.
    """
    # Valida o Workbook e a opcao recebida
    _validar_workbook(workbook)
    _validar_booleano(
        incluir_ocultas,
        "incluir_ocultas"
    )

    # Retorna todas as planilhas quando solicitado
    if incluir_ocultas:
        return list(workbook.sheetnames)

    # Retorna apenas as planilhas visiveis
    return [
        planilha.title
        for planilha in workbook.worksheets
        if planilha.sheet_state == "visible"
    ]


def contar_planilhas(
    workbook: Workbook,
    incluir_ocultas: bool = True
) -> int:
    """
    Retorna a quantidade de planilhas existentes no Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera consultado.
        incluir_ocultas (bool):
            Define se planilhas ocultas serao consideradas.

    Returns:
        int:
            Quantidade de planilhas encontradas.

    Raises:
        WorkbookPlanilhasInvalidoError:
            Caso o objeto nao seja um Workbook.
    """
    # Reutiliza a listagem centralizada das planilhas
    return len(
        listar_planilhas(
            workbook,
            incluir_ocultas=incluir_ocultas
        )
    )


def planilha_existe(
    workbook: Workbook,
    nome_planilha: str,
    considerar_maiusculas: bool = True
) -> bool:
    """
    Verifica se uma planilha existe no Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera consultado.
        nome_planilha (str):
            Nome da planilha procurada.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas e minusculas.

    Returns:
        bool:
            True quando a planilha existe; caso contrario, False.

    Raises:
        NomePlanilhaInvalidoError:
            Caso o nome informado seja invalido.
    """
    # Valida o Workbook, nome e opcao de comparacao
    _validar_workbook(workbook)
    nome = _validar_nome_planilha(nome_planilha)
    _validar_booleano(
        considerar_maiusculas,
        "considerar_maiusculas"
    )

    # Realiza a comparacao exata quando solicitado
    if considerar_maiusculas:
        return nome in workbook.sheetnames

    # Compara os nomes sem diferenciar caixa
    return nome.casefold() in {
        planilha.casefold()
        for planilha in workbook.sheetnames
    }


def obter_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet
) -> Worksheet:
    """
    Obtem uma planilha por nome, indice ou objeto Worksheet.

    Args:
        workbook (Workbook):
            Workbook que contem a planilha.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha.

    Returns:
        Worksheet:
            Planilha localizada no Workbook.

    Raises:
        PlanilhaNaoEncontradaError:
            Caso a planilha nao seja encontrada.
    """
    # Reutiliza a funcao interna de resolucao de referencias
    return _obter_planilha_por_referencia(
        workbook,
        referencia
    )


def obter_planilha_ativa(
    workbook: Workbook
) -> Worksheet:
    """
    Retorna a planilha ativa do Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera consultado.

    Returns:
        Worksheet:
            Planilha atualmente ativa.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso o Workbook nao possua planilhas.
    """
    # Valida o Workbook recebido
    _validar_workbook(workbook)

    # Verifica se existe ao menos uma planilha
    if not workbook.worksheets:
        raise OperacaoPlanilhaInvalidaError(
            "O Workbook nao possui planilhas."
        )

    # Retorna a planilha ativa
    return workbook.active


def obter_dimensoes_planilha(
    planilha: Worksheet
) -> dict[str, Any]:
    """
    Retorna as dimensoes e a regiao utilizada de uma planilha.

    Args:
        planilha (Worksheet):
            Planilha que sera consultada.

    Returns:
        dict[str, Any]:
            Dicionario com linhas, colunas, regiao e quantidade de celulas.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtem os limites informados pelo openpyxl
    linhas = planilha.max_row
    colunas = planilha.max_column
    dimensao = planilha.calculate_dimension()

    # Retorna as informacoes estruturais
    return {
        "quantidade_linhas": linhas,
        "quantidade_colunas": colunas,
        "quantidade_celulas": linhas * colunas,
        "intervalo_utilizado": dimensao,
        "ultima_coluna": get_column_letter(colunas)
    }


def obter_informacoes_planilha(
    planilha: Worksheet
) -> dict[str, Any]:
    """
    Retorna informacoes gerais de uma planilha Excel.

    Args:
        planilha (Worksheet):
            Planilha que sera consultada.

    Returns:
        dict[str, Any]:
            Dicionario com nome, indice, estado, dimensoes e configuracoes.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Valida a planilha
    _validar_planilha(planilha)

    # Obtem o Workbook pai e o indice da planilha
    workbook = planilha.parent
    indice = workbook.worksheets.index(planilha)

    # Retorna o resumo completo
    return {
        "nome": planilha.title,
        "indice": indice,
        "estado": planilha.sheet_state,
        "ativa": workbook.active is planilha,
        "cor_guia": (
            planilha.sheet_properties.tabColor.rgb
            if planilha.sheet_properties.tabColor is not None
            and planilha.sheet_properties.tabColor.type == "rgb"
            else None
        ),
        "protegida": planilha.protection.sheet,
        "congelamento": (
            str(planilha.freeze_panes)
            if planilha.freeze_panes is not None
            else None
        ),
        "linhas_grade": planilha.sheet_view.showGridLines,
        "filtro": planilha.auto_filter.ref,
        "tabelas": list(planilha.tables.keys()),
        "mesclagens": [
            str(intervalo)
            for intervalo in planilha.merged_cells.ranges
        ],
        **obter_dimensoes_planilha(planilha)
    }


# ----------------------------------------------------------------------------
# FUNCOES DE CRIACAO, RENOMEACAO E MOVIMENTACAO
# ----------------------------------------------------------------------------

def definir_planilha_ativa(
    workbook: Workbook,
    referencia: str | int | Worksheet
) -> Worksheet:
    """
    Define a planilha ativa do Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera alterado.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha.

    Returns:
        Worksheet:
            Planilha definida como ativa.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso a planilha esteja oculta.
    """
    # Obtem a planilha solicitada
    planilha = _obter_planilha_por_referencia(
        workbook,
        referencia
    )

    # O Excel nao permite ativar uma planilha oculta
    if planilha.sheet_state != "visible":
        raise OperacaoPlanilhaInvalidaError(
            "Uma planilha oculta nao pode ser definida como ativa."
        )

    # Define o indice ativo
    workbook.active = workbook.worksheets.index(planilha)

    # Retorna a planilha ativada
    return planilha


def criar_planilha(
    workbook: Workbook,
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    indice: int | None = None,
    ativar: bool = False,
    substituir_existente: bool = False
) -> Worksheet:
    """
    Cria uma nova planilha em um Workbook.

    Args:
        workbook (Workbook):
            Workbook que recebera a planilha.
        nome_planilha (str):
            Nome da planilha criada.
        indice (int | None):
            Posicao da planilha ou None para adicionar ao final.
        ativar (bool):
            Define se a nova planilha sera ativada.
        substituir_existente (bool):
            Define se uma planilha existente sera substituida.

    Returns:
        Worksheet:
            Planilha criada no Workbook.

    Raises:
        PlanilhaExistenteError:
            Caso o nome ja exista e a substituicao esteja desabilitada.
    """
    # Valida o Workbook, nome e opcoes
    _validar_workbook(workbook)
    nome = _validar_nome_planilha(nome_planilha)
    _validar_booleano(
        ativar,
        "ativar"
    )
    _validar_booleano(
        substituir_existente,
        "substituir_existente"
    )

    # Valida o indice quando informado
    if indice is not None:
        _validar_inteiro(
            indice,
            "indice",
            minimo=0,
            maximo=len(workbook.worksheets)
        )

    # Trata uma planilha existente
    if nome in workbook.sheetnames:
        if not substituir_existente:
            raise PlanilhaExistenteError(
                f"A planilha '{nome}' ja existe no Workbook."
            )

        existente = workbook[nome]

        # Impede a remocao da unica planilha sem reposicao imediata
        posicao_existente = workbook.worksheets.index(existente)
        workbook.remove(existente)

        if indice is None:
            indice = posicao_existente

    # Cria a nova planilha
    planilha = workbook.create_sheet(
        title=nome,
        index=indice
    )

    # Ativa a planilha quando solicitado
    if ativar:
        definir_planilha_ativa(
            workbook,
            planilha
        )

    # Retorna a planilha criada
    return planilha


def renomear_planilha(
    planilha: Worksheet,
    novo_nome: str
) -> Worksheet:
    """
    Renomeia uma planilha do Workbook.

    Args:
        planilha (Worksheet):
            Planilha que sera renomeada.
        novo_nome (str):
            Novo nome da planilha.

    Returns:
        Worksheet:
            A propria planilha renomeada.

    Raises:
        PlanilhaExistenteError:
            Caso outra planilha ja utilize o novo nome.
    """
    # Valida a planilha e o novo nome
    _validar_planilha(planilha)
    nome = _validar_nome_planilha(novo_nome)
    workbook = planilha.parent

    # Permite manter o nome atual
    if nome == planilha.title:
        return planilha

    # Verifica se o nome ja esta em uso
    if nome in workbook.sheetnames:
        raise PlanilhaExistenteError(
            f"A planilha '{nome}' ja existe no Workbook."
        )

    # Renomeia e retorna a planilha
    planilha.title = nome
    return planilha


def copiar_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet,
    novo_nome: str | None = None,
    ativar: bool = False
) -> Worksheet:
    """
    Copia uma planilha dentro do mesmo Workbook.

    Args:
        workbook (Workbook):
            Workbook que contem a planilha de origem.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha de origem.
        novo_nome (str | None):
            Nome opcional da copia.
        ativar (bool):
            Define se a copia sera ativada.

    Returns:
        Worksheet:
            Nova planilha copiada.

    Raises:
        PlanilhaExistenteError:
            Caso o novo nome ja esteja em uso.
    """
    # Valida o Workbook e localiza a origem
    _validar_workbook(workbook)
    origem = _obter_planilha_por_referencia(
        workbook,
        referencia
    )
    _validar_booleano(
        ativar,
        "ativar"
    )

    # Copia a planilha com o mecanismo nativo do openpyxl
    copia_planilha = workbook.copy_worksheet(origem)

    # Define o nome explicito quando informado
    if novo_nome is not None:
        nome = _validar_nome_planilha(novo_nome)

        if nome in workbook.sheetnames and workbook[nome] is not copia_planilha:
            workbook.remove(copia_planilha)
            raise PlanilhaExistenteError(
                f"A planilha '{nome}' ja existe no Workbook."
            )

        copia_planilha.title = nome

    # Ativa a copia quando solicitado
    if ativar:
        definir_planilha_ativa(
            workbook,
            copia_planilha
        )

    # Retorna a planilha copiada
    return copia_planilha


def copiar_planilha_entre_workbooks(
    planilha_origem: Worksheet,
    workbook_destino: Workbook,
    novo_nome: str | None = None
) -> Worksheet:
    """
    Copia valores, estilos e configuracoes entre Workbooks diferentes.

    Args:
        planilha_origem (Worksheet):
            Planilha utilizada como origem.
        workbook_destino (Workbook):
            Workbook que recebera a copia.
        novo_nome (str | None):
            Nome opcional da planilha copiada.

    Returns:
        Worksheet:
            Planilha criada no Workbook de destino.

    Raises:
        PlanilhaExistenteError:
            Caso o nome de destino ja exista.
    """
    # Valida a origem e o Workbook de destino
    _validar_planilha(planilha_origem)
    _validar_workbook(workbook_destino)

    # Define o nome da planilha copiada
    nome = _validar_nome_planilha(
        planilha_origem.title if novo_nome is None else novo_nome
    )

    # Rejeita um nome ja utilizado
    if nome in workbook_destino.sheetnames:
        raise PlanilhaExistenteError(
            f"A planilha '{nome}' ja existe no Workbook de destino."
        )

    # Cria a planilha de destino
    destino = workbook_destino.create_sheet(nome)

    # Copia valores e estilos das celulas
    for linha in planilha_origem.iter_rows():
        for celula_origem in linha:
            celula_destino = destino.cell(
                row=celula_origem.row,
                column=celula_origem.column,
                value=celula_origem.value
            )

            if celula_origem.has_style:
                celula_destino.font = copy(celula_origem.font)
                celula_destino.fill = copy(celula_origem.fill)
                celula_destino.border = copy(celula_origem.border)
                celula_destino.alignment = copy(celula_origem.alignment)
                celula_destino.number_format = celula_origem.number_format
                celula_destino.protection = copy(celula_origem.protection)

            if celula_origem.comment is not None:
                celula_destino.comment = copy(celula_origem.comment)

            if celula_origem.hyperlink is not None:
                celula_destino._hyperlink = copy(celula_origem.hyperlink)

    # Copia as dimensoes das colunas
    for letra, dimensao in planilha_origem.column_dimensions.items():
        destino.column_dimensions[letra].width = dimensao.width
        destino.column_dimensions[letra].hidden = dimensao.hidden
        destino.column_dimensions[letra].outlineLevel = dimensao.outlineLevel

    # Copia as dimensoes das linhas
    for indice, dimensao in planilha_origem.row_dimensions.items():
        destino.row_dimensions[indice].height = dimensao.height
        destino.row_dimensions[indice].hidden = dimensao.hidden
        destino.row_dimensions[indice].outlineLevel = dimensao.outlineLevel

    # Copia as mesclagens
    for intervalo in planilha_origem.merged_cells.ranges:
        destino.merge_cells(str(intervalo))

    # Copia configuracoes visuais e de impressao
    destino.freeze_panes = planilha_origem.freeze_panes
    destino.sheet_view.showGridLines = planilha_origem.sheet_view.showGridLines
    destino.auto_filter.ref = planilha_origem.auto_filter.ref
    destino.sheet_format.defaultColWidth = planilha_origem.sheet_format.defaultColWidth
    destino.sheet_format.defaultRowHeight = planilha_origem.sheet_format.defaultRowHeight
    destino.sheet_properties.tabColor = copy(
        planilha_origem.sheet_properties.tabColor
    )
    destino.sheet_properties.pageSetUpPr = copy(
        planilha_origem.sheet_properties.pageSetUpPr
    )
    destino.page_margins = copy(planilha_origem.page_margins)
    destino.page_setup = copy(planilha_origem.page_setup)
    destino.print_options = copy(planilha_origem.print_options)

    # Retorna a planilha copiada
    return destino


def mover_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet,
    indice_destino: int
) -> Worksheet:
    """
    Move uma planilha para uma posicao especifica do Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera reorganizado.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha movida.
        indice_destino (int):
            Novo indice da planilha, iniciado em zero.

    Returns:
        Worksheet:
            Planilha movida para a nova posicao.

    Raises:
        ValueError:
            Caso o indice de destino esteja fora dos limites.
    """
    # Localiza a planilha e valida o indice de destino
    planilha = _obter_planilha_por_referencia(
        workbook,
        referencia
    )
    destino = _validar_inteiro(
        indice_destino,
        "indice_destino",
        minimo=0,
        maximo=len(workbook.worksheets) - 1
    )

    # Obtem o indice atual
    atual = workbook.worksheets.index(planilha)

    # Move somente quando existe alteracao de posicao
    if atual != destino:
        workbook._sheets.pop(atual)
        workbook._sheets.insert(destino, planilha)

    # Retorna a planilha movida
    return planilha


def ordenar_planilhas(
    workbook: Workbook,
    ordem: Iterable[str] | None = None,
    alfabetica: bool = False,
    reversa: bool = False
) -> list[str]:
    """
    Ordena as planilhas de forma explicita ou alfabetica.

    Args:
        workbook (Workbook):
            Workbook que sera reorganizado.
        ordem (Iterable[str] | None):
            Ordem explicita dos nomes ou None para ordenacao alfabetica.
        alfabetica (bool):
            Define se os nomes serao ordenados alfabeticamente.
        reversa (bool):
            Define se a ordem final sera invertida.

    Returns:
        list[str]:
            Nomes das planilhas na nova ordem.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso a ordem explicita nao corresponda as planilhas existentes.
    """
    # Valida o Workbook e as opcoes
    _validar_workbook(workbook)
    _validar_booleano(
        alfabetica,
        "alfabetica"
    )
    _validar_booleano(
        reversa,
        "reversa"
    )

    # Impede duas estrategias simultaneas
    if ordem is not None and alfabetica:
        raise OperacaoPlanilhaInvalidaError(
            "Informe uma ordem explicita ou use alfabetica=True, nao ambos."
        )

    # Define a ordem alfabetica
    if alfabetica:
        nomes = sorted(
            workbook.sheetnames,
            key=str.casefold,
            reverse=reversa
        )

    # Define a ordem explicita
    elif ordem is not None:
        if isinstance(ordem, (str, bytes)) or not isinstance(ordem, Iterable):
            raise TypeError(
                "O parametro 'ordem' deve ser um iteravel nao textual ou None."
            )

        nomes = [
            _validar_nome_planilha(nome)
            for nome in ordem
        ]

        # Exige todos os nomes uma unica vez
        if len(nomes) != len(set(nomes)):
            raise OperacaoPlanilhaInvalidaError(
                "A ordem das planilhas nao pode possuir nomes duplicados."
            )

        if set(nomes) != set(workbook.sheetnames):
            raise OperacaoPlanilhaInvalidaError(
                "A ordem informada deve conter exatamente todas as planilhas."
            )

        if reversa:
            nomes.reverse()

    # Aplica apenas a inversao da ordem atual
    else:
        nomes = list(workbook.sheetnames)

        if reversa:
            nomes.reverse()
        else:
            return nomes

    # Reorganiza a lista interna de planilhas
    workbook._sheets = [
        workbook[nome]
        for nome in nomes
    ]

    # Retorna a nova ordem
    return nomes


# ----------------------------------------------------------------------------
# FUNCOES DE VISIBILIDADE, GUIA E PROTECAO
# ----------------------------------------------------------------------------

def ocultar_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet,
    muito_oculta: bool = False
) -> Worksheet:
    """
    Oculta uma planilha do Workbook.

    Args:
        workbook (Workbook):
            Workbook que contem a planilha.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha.
        muito_oculta (bool):
            Define se o estado veryHidden sera utilizado.

    Returns:
        Worksheet:
            Planilha ocultada.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso a operacao deixe o Workbook sem planilhas visiveis.
    """
    # Localiza a planilha e valida a opcao
    planilha = _obter_planilha_por_referencia(
        workbook,
        referencia
    )
    _validar_booleano(
        muito_oculta,
        "muito_oculta"
    )

    # Conta as demais planilhas visiveis
    visiveis_restantes = [
        item
        for item in workbook.worksheets
        if item is not planilha and item.sheet_state == "visible"
    ]

    # O Excel exige pelo menos uma planilha visivel
    if planilha.sheet_state == "visible" and not visiveis_restantes:
        raise OperacaoPlanilhaInvalidaError(
            "Nao e possivel ocultar a unica planilha visivel do Workbook."
        )

    # Ativa outra planilha antes de ocultar a ativa
    if workbook.active is planilha and visiveis_restantes:
        definir_planilha_ativa(
            workbook,
            visiveis_restantes[0]
        )

    # Define o estado de ocultacao
    planilha.sheet_state = "veryHidden" if muito_oculta else "hidden"

    # Retorna a planilha ocultada
    return planilha


def exibir_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet,
    ativar: bool = False
) -> Worksheet:
    """
    Torna uma planilha visivel novamente.

    Args:
        workbook (Workbook):
            Workbook que contem a planilha.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha.
        ativar (bool):
            Define se a planilha sera ativada apos ser exibida.

    Returns:
        Worksheet:
            Planilha tornada visivel.

    Raises:
        PlanilhaNaoEncontradaError:
            Caso a planilha nao seja encontrada.
    """
    # Localiza a planilha e valida a opcao
    planilha = _obter_planilha_por_referencia(
        workbook,
        referencia
    )
    _validar_booleano(
        ativar,
        "ativar"
    )

    # Torna a planilha visivel
    planilha.sheet_state = "visible"

    # Ativa a planilha quando solicitado
    if ativar:
        definir_planilha_ativa(
            workbook,
            planilha
        )

    # Retorna a planilha exibida
    return planilha


def definir_cor_guia(
    planilha: Worksheet,
    cor: str | None = COR_GUIA_PADRAO
) -> Worksheet:
    """
    Define ou remove a cor da guia de uma planilha.

    Args:
        planilha (Worksheet):
            Planilha cuja guia sera alterada.
        cor (str | None):
            Cor RGB ou ARGB, ou None para remover a cor.

    Returns:
        Worksheet:
            A propria planilha alterada.

    Raises:
        CorGuiaInvalidaError:
            Caso o codigo de cor seja invalido.
    """
    # Valida a planilha
    _validar_planilha(planilha)

    # Remove a cor quando None foi informado
    if cor is None:
        planilha.sheet_properties.tabColor = None
        return planilha

    # Normaliza e aplica a cor
    planilha.sheet_properties.tabColor = _normalizar_cor(cor)

    # Retorna a planilha alterada
    return planilha


def proteger_planilha(
    planilha: Worksheet,
    senha: str | None = None,
    selecionar_celulas_bloqueadas: bool = True,
    selecionar_celulas_desbloqueadas: bool = True
) -> Worksheet:
    """
    Ativa a protecao de uma planilha.

    Args:
        planilha (Worksheet):
            Planilha que sera protegida.
        senha (str | None):
            Senha opcional da protecao.
        selecionar_celulas_bloqueadas (bool):
            Define a selecao de celulas bloqueadas.
        selecionar_celulas_desbloqueadas (bool):
            Define a selecao de celulas desbloqueadas.

    Returns:
        Worksheet:
            A propria planilha protegida.

    Raises:
        TypeError:
            Caso a senha nao seja string ou None.
    """
    # Valida a planilha e as opcoes
    _validar_planilha(planilha)
    _validar_booleano(
        selecionar_celulas_bloqueadas,
        "selecionar_celulas_bloqueadas"
    )
    _validar_booleano(
        selecionar_celulas_desbloqueadas,
        "selecionar_celulas_desbloqueadas"
    )

    # Valida a senha quando informada
    if senha is not None:
        senha_validada = _validar_texto(
            senha,
            "senha"
        )
        planilha.protection.set_password(senha_validada)

    # Ativa e configura a protecao
    planilha.protection.sheet = True
    planilha.protection.selectLockedCells = selecionar_celulas_bloqueadas
    planilha.protection.selectUnlockedCells = selecionar_celulas_desbloqueadas

    # Retorna a planilha protegida
    return planilha


def desproteger_planilha(
    planilha: Worksheet
) -> Worksheet:
    """
    Desativa a protecao de uma planilha.

    Args:
        planilha (Worksheet):
            Planilha que sera desprotegida.

    Returns:
        Worksheet:
            A propria planilha desprotegida.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Valida a planilha
    _validar_planilha(planilha)

    # Desativa a protecao da planilha
    planilha.protection.sheet = False

    # Retorna a planilha desprotegida
    return planilha


# ----------------------------------------------------------------------------
# FUNCOES DE LIMPEZA E REMOCAO
# ----------------------------------------------------------------------------

def limpar_planilha(
    planilha: Worksheet,
    manter_formatacao: bool = True,
    manter_dimensoes: bool = True,
    manter_configuracoes: bool = True
) -> Worksheet:
    """
    Remove os conteudos de uma planilha e controla sua estrutura visual.

    Args:
        planilha (Worksheet):
            Planilha que sera limpa.
        manter_formatacao (bool):
            Define se os estilos das celulas serao preservados.
        manter_dimensoes (bool):
            Define se alturas e larguras serao preservadas.
        manter_configuracoes (bool):
            Define se filtros, mesclagens e congelamento serao preservados.

    Returns:
        Worksheet:
            A propria planilha limpa.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Valida a planilha e as opcoes
    _validar_planilha(planilha)

    for nome, opcao in {
        "manter_formatacao": manter_formatacao,
        "manter_dimensoes": manter_dimensoes,
        "manter_configuracoes": manter_configuracoes
    }.items():
        _validar_booleano(
            opcao,
            nome
        )

    # Limpa os valores ou remove integralmente as celulas
    if manter_formatacao:
        for linha in planilha.iter_rows():
            for celula in linha:
                # Ignora celulas mescladas secundarias, que sao somente leitura
                if isinstance(celula, MergedCell):
                    continue

                celula.value = None
                celula.comment = None
                celula.hyperlink = None
    else:
        for linha in planilha.iter_rows():
            for celula in linha:
                # Ignora celulas mescladas secundarias, que sao somente leitura
                if isinstance(celula, MergedCell):
                    continue

                celula.value = None
                celula.comment = None
                celula.hyperlink = None
                celula._style = StyleArray()
                celula.number_format = "General"

    # Remove as dimensoes personalizadas quando solicitado
    if not manter_dimensoes:
        planilha.column_dimensions.clear()
        planilha.row_dimensions.clear()

    # Remove configuracoes estruturais quando solicitado
    if not manter_configuracoes:
        planilha.freeze_panes = None
        planilha.auto_filter.ref = None
        planilha.merged_cells.ranges = []
        planilha.sheet_view.showGridLines = True
        planilha.data_validations.dataValidation = []

    # Retorna a planilha limpa
    return planilha


def remover_planilha(
    workbook: Workbook,
    referencia: str | int | Worksheet,
    permitir_workbook_vazio: bool = False
) -> str:
    """
    Remove uma planilha do Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera alterado.
        referencia (str | int | Worksheet):
            Nome, indice ou objeto da planilha removida.
        permitir_workbook_vazio (bool):
            Define se a ultima planilha podera ser removida.

    Returns:
        str:
            Nome da planilha removida.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso a ultima planilha nao possa ser removida.
    """
    # Localiza a planilha e valida a opcao
    planilha = _obter_planilha_por_referencia(
        workbook,
        referencia
    )
    _validar_booleano(
        permitir_workbook_vazio,
        "permitir_workbook_vazio"
    )

    # Protege a ultima planilha quando necessario
    if len(workbook.worksheets) == 1 and not permitir_workbook_vazio:
        raise OperacaoPlanilhaInvalidaError(
            "Nao e possivel remover a ultima planilha do Workbook."
        )

    # Armazena o nome e remove a planilha
    nome = planilha.title
    workbook.remove(planilha)

    # Retorna o nome removido
    return nome


def remover_planilhas(
    workbook: Workbook,
    referencias: Iterable[str | int | Worksheet],
    ignorar_ausentes: bool = False,
    permitir_workbook_vazio: bool = False
) -> list[str]:
    """
    Remove varias planilhas de um Workbook.

    Args:
        workbook (Workbook):
            Workbook que sera alterado.
        referencias (Iterable[str | int | Worksheet]):
            Colecao de nomes, indices ou objetos das planilhas.
        ignorar_ausentes (bool):
            Define se referencias inexistentes serao ignoradas.
        permitir_workbook_vazio (bool):
            Define se todas as planilhas poderao ser removidas.

    Returns:
        list[str]:
            Nomes das planilhas removidas.

    Raises:
        OperacaoPlanilhaInvalidaError:
            Caso a operacao deixe o Workbook vazio sem permissao.
    """
    # Valida o Workbook e as opcoes
    _validar_workbook(workbook)
    _validar_booleano(
        ignorar_ausentes,
        "ignorar_ausentes"
    )
    _validar_booleano(
        permitir_workbook_vazio,
        "permitir_workbook_vazio"
    )

    # Valida a colecao de referencias
    if isinstance(referencias, (str, bytes)) or not isinstance(
        referencias,
        Iterable
    ):
        raise TypeError(
            "O parametro 'referencias' deve ser um iteravel nao textual."
        )

    itens = list(referencias)

    if not itens:
        return []

    # Resolve as referencias antes de alterar o Workbook
    planilhas: list[Worksheet] = []

    for referencia in itens:
        try:
            planilha = _obter_planilha_por_referencia(
                workbook,
                referencia
            )
        except PlanilhaNaoEncontradaError:
            if ignorar_ausentes:
                continue
            raise

        if planilha not in planilhas:
            planilhas.append(planilha)

    # Verifica se a operacao deixaria o Workbook vazio
    if (
        len(planilhas) >= len(workbook.worksheets)
        and not permitir_workbook_vazio
    ):
        raise OperacaoPlanilhaInvalidaError(
            "A operacao removeria todas as planilhas do Workbook."
        )

    # Remove e registra os nomes das planilhas
    removidas: list[str] = []

    for planilha in planilhas:
        removidas.append(planilha.title)
        workbook.remove(planilha)

    # Retorna os nomes removidos
    return removidas


# ----------------------------------------------------------------------------
# EXPORTACOES PUBLICAS
# ----------------------------------------------------------------------------

__all__ = [
    "NOME_PLANILHA_PADRAO",
    "LIMITE_CARACTERES_PLANILHA",
    "CARACTERES_INVALIDOS_PLANILHA",
    "ESTADOS_VISIBILIDADE",
    "COR_GUIA_PADRAO",
    "ExcelPlanilhasError",
    "WorkbookPlanilhasInvalidoError",
    "PlanilhaInvalidaError",
    "NomePlanilhaInvalidoError",
    "PlanilhaNaoEncontradaError",
    "PlanilhaExistenteError",
    "OperacaoPlanilhaInvalidaError",
    "CorGuiaInvalidaError",
    "listar_planilhas",
    "contar_planilhas",
    "planilha_existe",
    "obter_planilha",
    "obter_planilha_ativa",
    "obter_dimensoes_planilha",
    "obter_informacoes_planilha",
    "definir_planilha_ativa",
    "criar_planilha",
    "renomear_planilha",
    "copiar_planilha",
    "copiar_planilha_entre_workbooks",
    "mover_planilha",
    "ordenar_planilhas",
    "ocultar_planilha",
    "exibir_planilha",
    "definir_cor_guia",
    "proteger_planilha",
    "desproteger_planilha",
    "limpar_planilha",
    "remover_planilha",
    "remover_planilhas"
]
