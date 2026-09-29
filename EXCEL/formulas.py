"""
============================================================
Módulo: EXCEL / formulas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para validar, definir, preencher,
    copiar, traduzir, localizar, remover e substituir fórmulas em
    planilhas Excel utilizando openpyxl.

    O openpyxl grava e preserva fórmulas, mas não realiza o cálculo dos
    resultados. O recálculo deve ser executado pelo Excel ou por outro
    mecanismo compatível após o arquivo ser aberto.
Funções disponíveis:
    - normalizar_formula()
    - validar_formula()
    - celula_possui_formula()
    - obter_formula()
    - definir_formula()
    - definir_formulas()
    - traduzir_formula()
    - copiar_formula()
    - preencher_formula_vertical()
    - preencher_formula_horizontal()
    - preencher_formula_intervalo()
    - listar_celulas_com_formula()
    - contar_formulas()
    - localizar_formulas()
    - substituir_texto_formulas()
    - remover_formula()
    - remover_formulas_intervalo()
    - converter_formula_em_valor()
    - converter_formulas_em_valores()
    - definir_modo_calculo()
    - forcar_recalculo()
Dependências:
    - openpyxl
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão das operações de leitura e escrita de fórmulas.
        - Inclusão de cópia e preenchimento com tradução de referências.
        - Inclusão das funções de pesquisa, substituição e remoção.
        - Inclusão das configurações de recálculo do Workbook.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa Any e Mapping para anotações flexíveis
from typing import Any, Mapping

# Importa Cell para validar objetos de célula
from openpyxl.cell.cell import Cell, MergedCell

# Importa Translator para ajustar referências relativas em fórmulas
from openpyxl.formula.translate import Translator

# Importa utilitários para validar coordenadas e intervalos
from openpyxl.utils import get_column_letter, range_boundaries

# Importa Workbook e Worksheet para validação dos objetos recebidos
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define o caractere obrigatório no início de uma fórmula Excel
PREFIXO_FORMULA = "="

# Define os modos de cálculo reconhecidos pelo Excel
MODOS_CALCULO = {
    "auto",
    "autoNoTable",
    "manual",
}

# Define o modo de cálculo padrão utilizado pela biblioteca
MODO_CALCULO_PADRAO = "auto"


# ----------------------------------------------------------------------------
# EXCEÇÕES
# ----------------------------------------------------------------------------

class ExcelFormulasError(Exception):
    """Erro base para operações relacionadas a fórmulas do Excel."""


class PlanilhaFormulaInvalidaError(ExcelFormulasError, TypeError):
    """Indica que o objeto informado não é uma Worksheet válida."""


class WorkbookFormulaInvalidoError(ExcelFormulasError, TypeError):
    """Indica que o objeto informado não é um Workbook válido."""


class ReferenciaCelulaInvalidaError(ExcelFormulasError, ValueError):
    """Indica que uma referência de célula ou intervalo é inválida."""


class FormulaInvalidaError(ExcelFormulasError, ValueError):
    """Indica que o texto informado não representa uma fórmula válida."""


class FormulaNaoEncontradaError(ExcelFormulasError, KeyError):
    """Indica que a célula consultada não contém uma fórmula."""


class TraducaoFormulaError(ExcelFormulasError):
    """Indica que uma fórmula não pôde ser traduzida para outro endereço."""


class ValorFormulaIndisponivelError(ExcelFormulasError):
    """Indica que o valor calculado de uma fórmula não está disponível."""


# ----------------------------------------------------------------------------
# FUNÇÕES INTERNAS DE VALIDAÇÃO
# ----------------------------------------------------------------------------

def _validar_planilha(planilha: Worksheet) -> Worksheet:
    """
    Valida se o objeto informado é uma planilha do openpyxl.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        PlanilhaFormulaInvalidaError:
            Caso o objeto informado não seja uma Worksheet válida.
    """

    # Verifica o tipo do objeto recebido
    if not isinstance(planilha, Worksheet):
        # Interrompe a execução para objetos incompatíveis
        raise PlanilhaFormulaInvalidaError(
            "O objeto fornecido não é uma planilha do openpyxl."
        )

    # Retorna a planilha validada
    return planilha


def _validar_workbook(workbook: Workbook) -> Workbook:
    """
    Valida se o objeto informado é um Workbook do openpyxl.

    Args:
        workbook (Workbook):
            Workbook do openpyxl que será validado ou configurado.

    Returns:
        Workbook:
            O próprio Workbook alterado ou validado.

    Raises:
        WorkbookFormulaInvalidoError:
            Caso o objeto informado não seja um Workbook válido.
    """

    # Verifica o tipo do objeto recebido
    if not isinstance(workbook, Workbook):
        # Interrompe a execução para objetos incompatíveis
        raise WorkbookFormulaInvalidoError(
            "O objeto fornecido não é um Workbook do openpyxl."
        )

    # Retorna o Workbook validado
    return workbook


def _validar_booleano(valor: bool, nome: str) -> bool:
    """
    Valida estritamente um parâmetro booleano.

    Args:
        valor (bool):
            Valor que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.

    Returns:
        bool:
            True quando a célula contém fórmula; caso contrário, False.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Verifica se o valor é realmente booleano
    if not isinstance(valor, bool):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser booleano."
        )

    # Retorna o valor validado
    return valor


def _validar_inteiro(
    valor: int,
    nome: str,
    minimo: int = 1
) -> int:
    """
    Valida um número inteiro e seu limite mínimo.

    Args:
        valor (int):
            Valor que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        minimo (int):
            Menor valor permitido para o parâmetro.

    Returns:
        int:
            Quantidade de fórmulas processadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Impede que booleanos sejam tratados como inteiros
    if isinstance(valor, bool) or not isinstance(valor, int):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser inteiro."
        )

    # Verifica o limite mínimo
    if valor < minimo:
        # Interrompe a execução quando o valor é menor que o permitido
        raise ValueError(
            f"O parâmetro '{nome}' deve ser maior ou igual a {minimo}."
        )

    # Retorna o inteiro validado
    return valor


def _validar_texto(valor: str, nome: str) -> str:
    """
    Valida e normaliza um parâmetro textual obrigatório.

    Args:
        valor (str):
            Valor que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Verifica se o valor é uma string
    if not isinstance(valor, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser uma string."
        )

    # Remove espaços externos
    texto = valor.strip()

    # Verifica se existe conteúdo útil
    if not texto:
        # Interrompe a execução quando o texto está vazio
        raise ValueError(
            f"O parâmetro '{nome}' não pode estar vazio."
        )

    # Retorna o texto normalizado
    return texto


def _normalizar_referencia_celula(referencia: str) -> str:
    """
    Valida e normaliza uma referência de célula única.

    Args:
        referencia (str):
            Referência da célula no padrão do Excel.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        ReferenciaCelulaInvalidaError:
            Caso a referência não represente uma única célula válida.
    """

    # Valida e converte a referência para letras maiúsculas
    coordenada = _validar_texto(
        referencia,
        "referencia"
    ).upper()

    # Tenta interpretar a coordenada como um intervalo
    try:
        coluna_inicial, linha_inicial, coluna_final, linha_final = (
            range_boundaries(coordenada)
        )
    except (TypeError, ValueError) as erro:
        # Converte a falha em uma exceção específica
        raise ReferenciaCelulaInvalidaError(
            f"A referência de célula é inválida: '{referencia}'."
        ) from erro

    # Agrupa os limites obtidos
    limites = (
        coluna_inicial,
        linha_inicial,
        coluna_final,
        linha_final
    )

    # Rejeita referências parciais como somente A ou somente 1
    if any(limite is None for limite in limites):
        # Interrompe a execução quando faltam linha ou coluna
        raise ReferenciaCelulaInvalidaError(
            f"A referência de célula é inválida: '{referencia}'."
        )

    # Rejeita índices de linha ou coluna menores que um
    if min(limites) < 1:
        raise ReferenciaCelulaInvalidaError(
            f"A referência de célula é inválida: '{referencia}'."
        )

    # Exige que o início e o fim representem a mesma célula
    if (
        coluna_inicial != coluna_final
        or linha_inicial != linha_final
    ):
        # Interrompe a execução quando foi informado um intervalo
        raise ReferenciaCelulaInvalidaError(
            "A referência deve representar uma única célula."
        )

    # Retorna a coordenada normalizada
    return coordenada


def _normalizar_referencia_intervalo(referencia: str) -> str:
    """
    Valida e normaliza uma referência de célula ou intervalo.

    Args:
        referencia (str):
            Referência da célula no padrão do Excel.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        ReferenciaCelulaInvalidaError:
            Caso a referência de célula ou intervalo seja inválida.
    """

    # Valida e converte a referência para letras maiúsculas
    intervalo = _validar_texto(
        referencia,
        "referencia"
    ).upper()

    # Tenta interpretar os limites do intervalo
    try:
        limites = range_boundaries(intervalo)
    except (TypeError, ValueError) as erro:
        # Converte a falha em uma exceção específica
        raise ReferenciaCelulaInvalidaError(
            f"A referência de intervalo é inválida: '{referencia}'."
        ) from erro

    # Rejeita referências parciais e limites nulos
    if any(limite is None for limite in limites):
        # Interrompe a execução quando linha ou coluna estão ausentes
        raise ReferenciaCelulaInvalidaError(
            f"A referência de intervalo é inválida: '{referencia}'."
        )

    # Rejeita índices de linha ou coluna menores que um
    if min(limites) < 1:
        raise ReferenciaCelulaInvalidaError(
            f"A referência de intervalo é inválida: '{referencia}'."
        )

    # Retorna o intervalo normalizado
    return intervalo


def _obter_celula(
    planilha: Worksheet,
    referencia: str
) -> Cell:
    """
    Obtém uma célula comum e rejeita células mescladas secundárias.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.

    Returns:
        Cell:
            Célula consultada ou alterada pela operação.

    Raises:
        ReferenciaCelulaInvalidaError:
            Caso a referência seja inválida ou pertença a uma célula mesclada secundária.
    """

    # Valida a planilha e a referência
    _validar_planilha(planilha)
    coordenada = _normalizar_referencia_celula(referencia)

    # Obtém a célula da planilha
    celula = planilha[coordenada]

    # Rejeita células secundárias pertencentes a uma mesclagem
    if isinstance(celula, MergedCell):
        # Interrompe a execução porque a célula não pode receber valores
        raise ReferenciaCelulaInvalidaError(
            f"A célula '{coordenada}' pertence a um intervalo mesclado."
        )

    # Retorna a célula válida
    return celula


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE VALIDAÇÃO E CONSULTA
# ----------------------------------------------------------------------------

def normalizar_formula(
    formula: str,
    adicionar_prefixo: bool = True
) -> str:
    """
    Normaliza um texto para o padrão de fórmula do Excel.

    Args:
        formula (str):
            Texto da fórmula do Excel.
        adicionar_prefixo (bool):
            Define se o caractere igual será adicionado quando estiver ausente.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o texto e a opção de inclusão do prefixo
    texto = _validar_texto(formula, "formula")
    _validar_booleano(adicionar_prefixo, "adicionar_prefixo")

    # Adiciona o prefixo quando solicitado e necessário
    if adicionar_prefixo and not texto.startswith(PREFIXO_FORMULA):
        texto = f"{PREFIXO_FORMULA}{texto}"

    # Retorna a fórmula normalizada
    return texto


def validar_formula(
    formula: str,
    exigir_prefixo: bool = True
) -> str:
    """
    Valida estruturalmente uma fórmula do Excel.

    Args:
        formula (str):
            Texto da fórmula do Excel.
        exigir_prefixo (bool):
            Define se o caractere igual será obrigatório no início da fórmula.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        FormulaInvalidaError:
            Caso o texto não represente uma fórmula estruturalmente válida.
    """

    # Valida o texto e a opção de prefixo obrigatório
    texto = _validar_texto(formula, "formula")
    _validar_booleano(exigir_prefixo, "exigir_prefixo")

    # Verifica o prefixo obrigatório
    if exigir_prefixo and not texto.startswith(PREFIXO_FORMULA):
        # Interrompe a execução quando o prefixo está ausente
        raise FormulaInvalidaError(
            "A fórmula deve começar com o caractere '='."
        )

    # Remove o prefixo apenas para validar o conteúdo
    conteudo = (
        texto[1:]
        if texto.startswith(PREFIXO_FORMULA)
        else texto
    ).strip()

    # Verifica se existe conteúdo após o prefixo
    if not conteudo:
        # Interrompe a execução quando a fórmula está vazia
        raise FormulaInvalidaError(
            "A fórmula deve possuir conteúdo após o caractere '='."
        )

    # Rejeita fórmulas com quebras de linha isoladas ou caracteres nulos
    if "\x00" in conteudo:
        # Interrompe a execução quando existe caractere nulo
        raise FormulaInvalidaError(
            "A fórmula possui caracteres inválidos."
        )

    # Retorna a fórmula validada
    return texto


def celula_possui_formula(
    planilha: Worksheet,
    referencia: str
) -> bool:
    """
    Verifica se uma célula contém uma fórmula.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.

    Returns:
        bool:
            True quando a célula contém fórmula; caso contrário, False.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Obtém a célula validada
    celula = _obter_celula(planilha, referencia)

    # Verifica o tipo interno ou o prefixo textual da célula
    return (
        celula.data_type == "f"
        or (
            isinstance(celula.value, str)
            and celula.value.startswith(PREFIXO_FORMULA)
        )
    )


def obter_formula(
    planilha: Worksheet,
    referencia: str,
    permitir_ausente: bool = False
) -> str | None:
    """
    Obtém a fórmula armazenada em uma célula.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.
        permitir_ausente (bool):
            Define se a ausência de fórmula poderá retornar None.

    Returns:
        str | None:
            Fórmula encontrada ou None quando a ausência for permitida.

    Raises:
        FormulaNaoEncontradaError:
            Caso a célula não contenha fórmula e a ausência não seja permitida.
    """

    # Valida a opção de retorno ausente
    _validar_booleano(permitir_ausente, "permitir_ausente")

    # Obtém a célula validada
    celula = _obter_celula(planilha, referencia)

    # Retorna a fórmula quando ela existe
    if celula_possui_formula(planilha, referencia):
        return str(celula.value)

    # Retorna None quando a ausência foi permitida
    if permitir_ausente:
        return None

    # Interrompe a execução quando a célula não contém fórmula
    raise FormulaNaoEncontradaError(
        f"A célula '{celula.coordinate}' não contém uma fórmula."
    )


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE DEFINIÇÃO E CÓPIA
# ----------------------------------------------------------------------------

def definir_formula(
    planilha: Worksheet,
    referencia: str,
    formula: str,
    adicionar_prefixo: bool = True
) -> Cell:
    """
    Define uma fórmula em uma célula.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.
        formula (str):
            Texto da fórmula do Excel.
        adicionar_prefixo (bool):
            Define se o caractere igual será adicionado quando estiver ausente.

    Returns:
        Cell:
            Célula consultada ou alterada pela operação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Obtém a célula de destino
    celula = _obter_celula(planilha, referencia)

    # Normaliza e valida a fórmula
    formula_normalizada = normalizar_formula(
        formula,
        adicionar_prefixo=adicionar_prefixo
    )
    validar_formula(formula_normalizada)

    # Grava a fórmula na célula
    celula.value = formula_normalizada

    # Retorna a célula alterada
    return celula


def definir_formulas(
    planilha: Worksheet,
    formulas: Mapping[str, str],
    adicionar_prefixo: bool = True
) -> Worksheet:
    """
    Define várias fórmulas a partir de um mapeamento de células.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        formulas (Mapping[str, str]):
            Mapeamento entre referências de células e fórmulas.
        adicionar_prefixo (bool):
            Define se o caractere igual será adicionado quando estiver ausente.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a opção de prefixo
    _validar_planilha(planilha)
    _validar_booleano(adicionar_prefixo, "adicionar_prefixo")

    # Verifica se as fórmulas foram informadas como mapeamento
    if not isinstance(formulas, Mapping):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'formulas' deve ser um mapeamento."
        )

    # Verifica se o mapeamento possui conteúdo
    if not formulas:
        # Interrompe a execução quando nenhuma fórmula foi informada
        raise ValueError(
            "O parâmetro 'formulas' não pode estar vazio."
        )

    # Percorre as referências e fórmulas na ordem recebida
    for referencia, formula in formulas.items():
        # Define cada fórmula na célula correspondente
        definir_formula(
            planilha,
            referencia,
            formula,
            adicionar_prefixo=adicionar_prefixo
        )

    # Retorna a planilha alterada
    return planilha


def traduzir_formula(
    formula: str,
    origem: str,
    destino: str
) -> str:
    """
    Traduz as referências relativas de uma fórmula para outra célula.

    Args:
        formula (str):
            Texto da fórmula do Excel.
        origem (str):
            Referência da célula de origem da fórmula.
        destino (str):
            Referência da célula de destino da fórmula.

    Returns:
        str:
            Texto normalizado, validado ou traduzido.

    Raises:
        TraducaoFormulaError:
            Caso as referências da fórmula não possam ser traduzidas.
    """

    # Valida a fórmula e as referências envolvidas
    formula_validada = validar_formula(
        normalizar_formula(formula)
    )
    origem_validada = _normalizar_referencia_celula(origem)
    destino_validado = _normalizar_referencia_celula(destino)

    # Executa a tradução através do openpyxl
    try:
        formula_traduzida = Translator(
            formula_validada,
            origin=origem_validada
        ).translate_formula(destino_validado)
    except Exception as erro:
        # Converte a falha em uma exceção específica do módulo
        raise TraducaoFormulaError(
            f"Não foi possível traduzir a fórmula de "
            f"'{origem_validada}' para '{destino_validado}'."
        ) from erro

    # Retorna a fórmula traduzida
    return formula_traduzida


def copiar_formula(
    planilha: Worksheet,
    origem: str,
    destino: str,
    traduzir_referencias: bool = True
) -> Cell:
    """
    Copia uma fórmula entre células.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        origem (str):
            Referência da célula de origem da fórmula.
        destino (str):
            Referência da célula de destino da fórmula.
        traduzir_referencias (bool):
            Define se as referências relativas serão ajustadas para o destino.

    Returns:
        Cell:
            Célula consultada ou alterada pela operação.

    Raises:
        FormulaNaoEncontradaError:
            Caso a célula de origem não contenha fórmula.
    """

    # Valida a planilha e a opção de tradução
    _validar_planilha(planilha)
    _validar_booleano(
        traduzir_referencias,
        "traduzir_referencias"
    )

    # Obtém a fórmula da célula de origem
    formula_origem = obter_formula(planilha, origem)

    # Define a fórmula final com ou sem tradução
    formula_destino = (
        traduzir_formula(
            formula_origem,
            origem,
            destino
        )
        if traduzir_referencias
        else formula_origem
    )

    # Grava a fórmula na célula de destino
    return definir_formula(
        planilha,
        destino,
        formula_destino,
        adicionar_prefixo=False
    )


def preencher_formula_vertical(
    planilha: Worksheet,
    referencia_origem: str,
    linha_final: int,
    traduzir_referencias: bool = True
) -> Worksheet:
    """
    Preenche uma fórmula verticalmente até a linha informada.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia_origem (str):
            Referência da célula que contém a fórmula de origem.
        linha_final (int):
            Última linha que receberá a fórmula.
        traduzir_referencias (bool):
            Define se as referências relativas serão ajustadas para o destino.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha, a origem e a linha final
    _validar_planilha(planilha)
    origem = _normalizar_referencia_celula(referencia_origem)
    celula_origem = _obter_celula(planilha, origem)
    final = _validar_inteiro(
        linha_final,
        "linha_final",
        minimo=celula_origem.row
    )

    # Valida a opção de tradução
    _validar_booleano(
        traduzir_referencias,
        "traduzir_referencias"
    )

    # Garante que a origem possui uma fórmula
    obter_formula(planilha, origem)

    # Percorre as linhas posteriores à origem
    for linha in range(celula_origem.row + 1, final + 1):
        # Monta a referência de destino na mesma coluna
        destino = f"{get_column_letter(celula_origem.column)}{linha}"

        # Copia a fórmula para a célula de destino
        copiar_formula(
            planilha,
            origem,
            destino,
            traduzir_referencias=traduzir_referencias
        )

    # Retorna a planilha alterada
    return planilha


def preencher_formula_horizontal(
    planilha: Worksheet,
    referencia_origem: str,
    coluna_final: int | str,
    traduzir_referencias: bool = True
) -> Worksheet:
    """
    Preenche uma fórmula horizontalmente até a coluna informada.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia_origem (str):
            Referência da célula que contém a fórmula de origem.
        coluna_final (int | str):
            Última coluna que receberá a fórmula.
        traduzir_referencias (bool):
            Define se as referências relativas serão ajustadas para o destino.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e obtém a célula de origem
    _validar_planilha(planilha)
    origem = _normalizar_referencia_celula(referencia_origem)
    celula_origem = _obter_celula(planilha, origem)

    # Normaliza a referência da última coluna
    if isinstance(coluna_final, int) and not isinstance(coluna_final, bool):
        indice_final = _validar_inteiro(
            coluna_final,
            "coluna_final",
            minimo=celula_origem.column
        )
    elif isinstance(coluna_final, str):
        # Converte a letra em uma coordenada completa para validação
        referencia_final = _normalizar_referencia_celula(
            f"{_validar_texto(coluna_final, 'coluna_final').upper()}1"
        )
        indice_final = range_boundaries(referencia_final)[0]

        # Verifica se a última coluna não está antes da origem
        if indice_final < celula_origem.column:
            raise ValueError(
                "A coluna final não pode estar antes da coluna de origem."
            )
    else:
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'coluna_final' deve ser string ou inteiro."
        )

    # Valida a opção de tradução e a fórmula de origem
    _validar_booleano(
        traduzir_referencias,
        "traduzir_referencias"
    )
    obter_formula(planilha, origem)

    # Percorre as colunas posteriores à origem
    for coluna in range(celula_origem.column + 1, indice_final + 1):
        # Monta a referência de destino na mesma linha
        destino = f"{get_column_letter(coluna)}{celula_origem.row}"

        # Copia a fórmula para o destino
        copiar_formula(
            planilha,
            origem,
            destino,
            traduzir_referencias=traduzir_referencias
        )

    # Retorna a planilha alterada
    return planilha


def preencher_formula_intervalo(
    planilha: Worksheet,
    referencia_origem: str,
    intervalo_destino: str,
    traduzir_referencias: bool = True,
    incluir_origem: bool = False
) -> Worksheet:
    """
    Preenche uma fórmula em todas as células de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia_origem (str):
            Referência da célula que contém a fórmula de origem.
        intervalo_destino (str):
            Intervalo que receberá a fórmula de origem.
        traduzir_referencias (bool):
            Define se as referências relativas serão ajustadas para o destino.
        incluir_origem (bool):
            Define se a célula de origem será incluída no preenchimento.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha, origem, intervalo e opções
    _validar_planilha(planilha)
    origem = _normalizar_referencia_celula(referencia_origem)
    intervalo = _normalizar_referencia_intervalo(intervalo_destino)
    _validar_booleano(
        traduzir_referencias,
        "traduzir_referencias"
    )
    _validar_booleano(incluir_origem, "incluir_origem")

    # Garante que a origem contém uma fórmula
    formula_origem = obter_formula(planilha, origem)

    # Obtém os limites do intervalo de destino
    coluna_inicial, linha_inicial, coluna_final, linha_final = (
        range_boundaries(intervalo)
    )

    # Percorre todas as células do intervalo
    for linha in range(linha_inicial, linha_final + 1):
        for coluna in range(coluna_inicial, coluna_final + 1):
            # Monta a coordenada da célula atual
            destino = f"{get_column_letter(coluna)}{linha}"

            # Ignora a origem quando solicitado
            if destino == origem and not incluir_origem:
                continue

            # Traduz ou reutiliza a fórmula original
            formula_destino = (
                traduzir_formula(
                    formula_origem,
                    origem,
                    destino
                )
                if traduzir_referencias
                else formula_origem
            )

            # Define a fórmula na célula atual
            definir_formula(
                planilha,
                destino,
                formula_destino,
                adicionar_prefixo=False
            )

    # Retorna a planilha alterada
    return planilha


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE PESQUISA E SUBSTITUIÇÃO
# ----------------------------------------------------------------------------

def listar_celulas_com_formula(
    planilha: Worksheet,
    intervalo: str | None = None
) -> list[dict[str, Any]]:
    """
    Lista as células que contêm fórmulas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        intervalo (str | None):
            Intervalo opcional utilizado para limitar a operação.

    Returns:
        list[dict[str, Any]]:
            Lista com os dados das fórmulas encontradas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Define os limites da pesquisa
    if intervalo is None:
        coluna_inicial = 1
        linha_inicial = 1
        coluna_final = planilha.max_column
        linha_final = planilha.max_row
    else:
        # Normaliza e converte o intervalo informado
        referencia = _normalizar_referencia_intervalo(intervalo)
        coluna_inicial, linha_inicial, coluna_final, linha_final = (
            range_boundaries(referencia)
        )

    # Cria a lista de fórmulas localizadas
    resultado: list[dict[str, Any]] = []

    # Percorre as células do intervalo
    for linha in range(linha_inicial, linha_final + 1):
        for coluna in range(coluna_inicial, coluna_final + 1):
            # Obtém a célula atual
            celula = planilha.cell(
                row=linha,
                column=coluna
            )

            # Verifica se a célula contém fórmula
            if (
                celula.data_type == "f"
                or (
                    isinstance(celula.value, str)
                    and celula.value.startswith(PREFIXO_FORMULA)
                )
            ):
                # Adiciona os dados da fórmula ao resultado
                resultado.append(
                    {
                        "coordenada": celula.coordinate,
                        "formula": str(celula.value),
                        "linha": linha,
                        "coluna": coluna,
                    }
                )

    # Retorna as fórmulas encontradas
    return resultado


def contar_formulas(
    planilha: Worksheet,
    intervalo: str | None = None
) -> int:
    """
    Retorna a quantidade de fórmulas encontradas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        intervalo (str | None):
            Intervalo opcional utilizado para limitar a operação.

    Returns:
        int:
            Quantidade de fórmulas processadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Reutiliza a listagem centralizada de fórmulas
    return len(
        listar_celulas_com_formula(
            planilha,
            intervalo=intervalo
        )
    )


def localizar_formulas(
    planilha: Worksheet,
    texto: str,
    intervalo: str | None = None,
    considerar_maiusculas: bool = False
) -> list[dict[str, Any]]:
    """
    Localiza fórmulas que contêm um texto específico.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        texto (str):
            Texto que será procurado nas fórmulas.
        intervalo (str | None):
            Intervalo opcional utilizado para limitar a operação.
        considerar_maiusculas (bool):
            Define se a comparação diferencia letras maiúsculas e minúsculas.

    Returns:
        list[dict[str, Any]]:
            Lista com os dados das fórmulas encontradas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o texto pesquisado e a opção de comparação
    texto_procurado = _validar_texto(texto, "texto")
    _validar_booleano(
        considerar_maiusculas,
        "considerar_maiusculas"
    )

    # Normaliza o texto quando a busca não diferencia maiúsculas
    esperado = (
        texto_procurado
        if considerar_maiusculas
        else texto_procurado.casefold()
    )

    # Cria a lista dos resultados correspondentes
    resultado: list[dict[str, Any]] = []

    # Percorre as fórmulas disponíveis
    for dados in listar_celulas_com_formula(
        planilha,
        intervalo=intervalo
    ):
        # Normaliza a fórmula atual conforme a configuração
        formula_atual = (
            dados["formula"]
            if considerar_maiusculas
            else dados["formula"].casefold()
        )

        # Adiciona as fórmulas que contêm o texto procurado
        if esperado in formula_atual:
            resultado.append(dados)

    # Retorna as fórmulas localizadas
    return resultado


def substituir_texto_formulas(
    planilha: Worksheet,
    texto_atual: str,
    novo_texto: str,
    intervalo: str | None = None,
    considerar_maiusculas: bool = True
) -> int:
    """
    Substitui um texto nas fórmulas da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        texto_atual (str):
            Texto atual que será localizado nas fórmulas.
        novo_texto (str):
            Novo texto que substituirá as ocorrências encontradas.
        intervalo (str | None):
            Intervalo opcional utilizado para limitar a operação.
        considerar_maiusculas (bool):
            Define se a comparação diferencia letras maiúsculas e minúsculas.

    Returns:
        int:
            Quantidade de fórmulas processadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida os textos e a opção de comparação
    atual = _validar_texto(texto_atual, "texto_atual")
    if not isinstance(novo_texto, str):
        raise TypeError(
            "O parâmetro 'novo_texto' deve ser uma string."
        )
    _validar_booleano(
        considerar_maiusculas,
        "considerar_maiusculas"
    )

    # Inicia o contador de fórmulas alteradas
    quantidade = 0

    # Percorre todas as fórmulas encontradas
    for dados in listar_celulas_com_formula(
        planilha,
        intervalo=intervalo
    ):
        # Obtém a fórmula atual
        formula = dados["formula"]

        # Realiza a substituição direta quando a comparação diferencia caixa
        if considerar_maiusculas:
            nova_formula = formula.replace(atual, novo_texto)
        else:
            # Localiza e substitui ocorrências preservando o restante do texto
            nova_formula = formula
            inicio_busca = 0
            while True:
                indice = nova_formula.casefold().find(
                    atual.casefold(),
                    inicio_busca
                )
                if indice < 0:
                    break
                nova_formula = (
                    nova_formula[:indice]
                    + novo_texto
                    + nova_formula[indice + len(atual):]
                )
                inicio_busca = indice + len(novo_texto)

        # Atualiza somente as fórmulas que foram modificadas
        if nova_formula != formula:
            definir_formula(
                planilha,
                dados["coordenada"],
                nova_formula,
                adicionar_prefixo=False
            )
            quantidade += 1

    # Retorna a quantidade de fórmulas alteradas
    return quantidade


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE REMOÇÃO E CONVERSÃO
# ----------------------------------------------------------------------------

def remover_formula(
    planilha: Worksheet,
    referencia: str,
    valor_substituto: Any = None,
    exigir_formula: bool = True
) -> Cell:
    """
    Remove uma fórmula e grava um valor substituto.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.
        valor_substituto (Any):
            Valor que substituirá a fórmula removida.
        exigir_formula (bool):
            Define se a célula deve necessariamente conter uma fórmula.

    Returns:
        Cell:
            Célula consultada ou alterada pela operação.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a opção que exige fórmula
    _validar_booleano(exigir_formula, "exigir_formula")

    # Obtém a célula de destino
    celula = _obter_celula(planilha, referencia)

    # Verifica se a célula contém fórmula quando isso é obrigatório
    if exigir_formula and not celula_possui_formula(planilha, referencia):
        # Interrompe a execução quando nenhuma fórmula foi encontrada
        raise FormulaNaoEncontradaError(
            f"A célula '{celula.coordinate}' não contém uma fórmula."
        )

    # Substitui a fórmula pelo valor informado
    celula.value = valor_substituto

    # Retorna a célula alterada
    return celula


def remover_formulas_intervalo(
    planilha: Worksheet,
    intervalo: str,
    valor_substituto: Any = None
) -> int:
    """
    Remove todas as fórmulas de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        intervalo (str):
            Intervalo opcional utilizado para limitar a operação.
        valor_substituto (Any):
            Valor que substituirá a fórmula removida.

    Returns:
        int:
            Quantidade de fórmulas processadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Lista as fórmulas existentes no intervalo
    formulas = listar_celulas_com_formula(
        planilha,
        intervalo=intervalo
    )

    # Percorre as fórmulas encontradas
    for dados in formulas:
        # Remove cada fórmula sem uma nova validação obrigatória
        remover_formula(
            planilha,
            dados["coordenada"],
            valor_substituto=valor_substituto,
            exigir_formula=False
        )

    # Retorna a quantidade removida
    return len(formulas)


def converter_formula_em_valor(
    planilha: Worksheet,
    referencia: str,
    valor_calculado: Any,
    permitir_none: bool = False
) -> Cell:
    """
    Substitui uma fórmula por um valor previamente calculado.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula no padrão do Excel.
        valor_calculado (Any):
            Valor previamente calculado que substituirá a fórmula.
        permitir_none (bool):
            Define se None poderá substituir uma fórmula.

    Returns:
        Cell:
            Célula consultada ou alterada pela operação.

    Raises:
        ValorFormulaIndisponivelError:
            Caso o valor calculado seja None e isso não seja permitido.
    """

    # Valida a opção de valores ausentes
    _validar_booleano(permitir_none, "permitir_none")

    # Garante que a célula contém uma fórmula
    obter_formula(planilha, referencia)

    # Rejeita valores ausentes quando não permitidos
    if valor_calculado is None and not permitir_none:
        # Interrompe a execução para evitar perda acidental da fórmula
        raise ValorFormulaIndisponivelError(
            "O valor calculado da fórmula não está disponível."
        )

    # Remove a fórmula e mantém o valor calculado
    return remover_formula(
        planilha,
        referencia,
        valor_substituto=valor_calculado,
        exigir_formula=False
    )


def converter_formulas_em_valores(
    planilha_formulas: Worksheet,
    planilha_valores: Worksheet,
    intervalo: str | None = None,
    permitir_none: bool = False
) -> int:
    """
    Substitui fórmulas usando valores de outra planilha equivalente.

    Args:
        planilha_formulas (Worksheet):
            Planilha que contém as fórmulas originais.
        planilha_valores (Worksheet):
            Planilha equivalente que contém os valores calculados.
        intervalo (str | None):
            Intervalo opcional utilizado para limitar a operação.
        permitir_none (bool):
            Define se None poderá substituir uma fórmula.

    Returns:
        int:
            Quantidade de fórmulas processadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida as duas planilhas e a opção de valores ausentes
    _validar_planilha(planilha_formulas)
    _validar_planilha(planilha_valores)
    _validar_booleano(permitir_none, "permitir_none")

    # Verifica se as duas planilhas representam a mesma aba lógica
    if planilha_formulas.title != planilha_valores.title:
        # Interrompe a execução para evitar cruzamento incorreto de células
        raise ValueError(
            "As planilhas de fórmulas e valores devem possuir o mesmo título."
        )

    # Lista as fórmulas que serão convertidas
    formulas = listar_celulas_com_formula(
        planilha_formulas,
        intervalo=intervalo
    )

    # Percorre as fórmulas localizadas
    for dados in formulas:
        # Obtém o valor calculado da planilha paralela
        valor = planilha_valores[dados["coordenada"]].value

        # Converte a fórmula mantendo o valor correspondente
        converter_formula_em_valor(
            planilha_formulas,
            dados["coordenada"],
            valor,
            permitir_none=permitir_none
        )

    # Retorna a quantidade de fórmulas convertidas
    return len(formulas)


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE CÁLCULO DO WORKBOOK
# ----------------------------------------------------------------------------

def definir_modo_calculo(
    workbook: Workbook,
    modo: str = MODO_CALCULO_PADRAO
) -> Workbook:
    """
    Define o modo de cálculo armazenado no Workbook.

    Args:
        workbook (Workbook):
            Workbook do openpyxl que será validado ou configurado.
        modo (str):
            Modo de cálculo que será configurado no Workbook.

    Returns:
        Workbook:
            O próprio Workbook alterado ou validado.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o Workbook e normaliza o modo recebido
    _validar_workbook(workbook)
    modo_validado = _validar_texto(modo, "modo")

    # Verifica se o modo é reconhecido pelo Excel
    if modo_validado not in MODOS_CALCULO:
        # Interrompe a execução quando o modo é inválido
        raise ValueError(
            f"Modo de cálculo inválido: '{modo_validado}'."
        )

    # Define o modo de cálculo nas propriedades do Workbook
    workbook.calculation.calcMode = modo_validado

    # Retorna o Workbook alterado
    return workbook


def forcar_recalculo(
    workbook: Workbook,
    recalculo_completo: bool = True
) -> Workbook:
    """
    Marca o Workbook para recálculo na próxima abertura pelo Excel.

    Args:
        workbook (Workbook):
            Workbook do openpyxl que será validado ou configurado.
        recalculo_completo (bool):
            Define se o Excel deverá executar um recálculo completo.

    Returns:
        Workbook:
            O próprio Workbook alterado ou validado.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o Workbook e a opção de recálculo completo
    _validar_workbook(workbook)
    _validar_booleano(
        recalculo_completo,
        "recalculo_completo"
    )

    # Habilita o modo automático de cálculo
    workbook.calculation.calcMode = "auto"

    # Solicita cálculo ao abrir o arquivo
    workbook.calculation.calcOnSave = True
    workbook.calculation.forceFullCalc = recalculo_completo
    workbook.calculation.fullCalcOnLoad = recalculo_completo

    # Retorna o Workbook alterado
    return workbook


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

# Define explicitamente os objetos públicos disponibilizados pelo módulo
__all__ = [
    "PREFIXO_FORMULA",
    "MODOS_CALCULO",
    "MODO_CALCULO_PADRAO",
    "ExcelFormulasError",
    "PlanilhaFormulaInvalidaError",
    "WorkbookFormulaInvalidoError",
    "ReferenciaCelulaInvalidaError",
    "FormulaInvalidaError",
    "FormulaNaoEncontradaError",
    "TraducaoFormulaError",
    "ValorFormulaIndisponivelError",
    "normalizar_formula",
    "validar_formula",
    "celula_possui_formula",
    "obter_formula",
    "definir_formula",
    "definir_formulas",
    "traduzir_formula",
    "copiar_formula",
    "preencher_formula_vertical",
    "preencher_formula_horizontal",
    "preencher_formula_intervalo",
    "listar_celulas_com_formula",
    "contar_formulas",
    "localizar_formulas",
    "substituir_texto_formulas",
    "remover_formula",
    "remover_formulas_intervalo",
    "converter_formula_em_valor",
    "converter_formulas_em_valores",
    "definir_modo_calculo",
    "forcar_recalculo",
]
