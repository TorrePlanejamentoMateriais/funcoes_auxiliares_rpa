"""
============================================================
Módulo: EXCEL / formatacao.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para aplicar fontes, preenchimentos,
    bordas, alinhamentos, formatos numéricos, estilos de cabeçalho,
    dimensões, congelamento de painéis e configurações visuais em
    planilhas Excel utilizando openpyxl.
Funções disponíveis:
    - normalizar_cor()
    - obter_celulas_intervalo()
    - aplicar_fonte()
    - aplicar_preenchimento()
    - aplicar_alinhamento()
    - aplicar_bordas()
    - aplicar_formato_numero()
    - aplicar_formato_data()
    - aplicar_formato_moeda()
    - aplicar_formato_percentual()
    - formatar_cabecalho()
    - formatar_intervalo()
    - copiar_formatacao()
    - limpar_formatacao()
    - definir_largura_coluna()
    - definir_altura_linha()
    - ajustar_larguras_colunas()
    - congelar_paineis()
    - descongelar_paineis()
    - mesclar_celulas()
    - desmesclar_celulas()
    - ocultar_linhas_grade()
    - exibir_linhas_grade()
    - aplicar_estilo_zebrado()
Dependências:
    - openpyxl
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão das funções de fonte, preenchimento e alinhamento.
        - Inclusão das funções de bordas e formatos numéricos.
        - Inclusão da formatação de cabeçalhos e intervalos.
        - Inclusão das configurações de dimensões e visualização.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa copy para duplicar estilos sem compartilhar objetos mutáveis
from copy import copy

# Importa tipos auxiliares utilizados nas anotações das funções
from typing import Any, Iterable

# Importa as classes de estilo disponibilizadas pelo openpyxl
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Protection,
    Side
)

# Importa a classe Cell para validar células e restaurar estilos padrão
from openpyxl.cell.cell import Cell

# Importa utilitários para manipulação de referências e colunas
from openpyxl.utils import get_column_letter, range_boundaries

# Importa a classe Worksheet para validar as planilhas recebidas
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define as cores padrão utilizadas nos cabeçalhos
COR_CABECALHO_PADRAO = "FF1F4E78"
COR_TEXTO_CABECALHO_PADRAO = "FFFFFFFF"

# Define cores auxiliares utilizadas em estilos alternados
COR_LINHA_ALTERNADA_PADRAO = "FFF2F2F2"
COR_BORDA_PADRAO = "FFD9D9D9"

# Define os formatos numéricos padrão
FORMATO_DATA_PADRAO = "dd/mm/yyyy"
FORMATO_MOEDA_PADRAO = 'R$ #,##0.00;[Red](R$ #,##0.00)'
FORMATO_PERCENTUAL_PADRAO = "0.0%"
FORMATO_NUMERO_PADRAO = "#,##0.00"

# Define os limites práticos das dimensões do Excel
LARGURA_MAXIMA_COLUNA = 255
ALTURA_MAXIMA_LINHA = 409

# Define os estilos aceitos para bordas
TIPOS_BORDA = {
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
}

# Define os alinhamentos aceitos pelo openpyxl
ALINHAMENTOS_HORIZONTAIS = {
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
}
ALINHAMENTOS_VERTICAIS = {
    "top",
    "center",
    "bottom",
    "justify",
    "distributed",
}


# ----------------------------------------------------------------------------
# EXCEÇÕES
# ----------------------------------------------------------------------------

class ExcelFormatacaoError(Exception):
    """Erro base para operações de formatação de planilhas Excel."""


class PlanilhaFormatacaoInvalidaError(ExcelFormatacaoError, TypeError):
    """Indica que o objeto informado não é uma Worksheet válida."""


class IntervaloFormatacaoInvalidoError(ExcelFormatacaoError, ValueError):
    """Indica que a referência do intervalo é inválida."""


class CorInvalidaError(ExcelFormatacaoError, ValueError):
    """Indica que o código de cor não possui formato hexadecimal válido."""


class EstiloInvalidoError(ExcelFormatacaoError, ValueError):
    """Indica que uma configuração de estilo não é suportada."""


class DimensaoInvalidaError(ExcelFormatacaoError, ValueError):
    """Indica que uma largura ou altura está fora dos limites permitidos."""


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
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Verifica se o objeto recebido é uma planilha válida
    if not isinstance(planilha, Worksheet):
        # Interrompe a execução quando o tipo é inválido
        raise PlanilhaFormatacaoInvalidaError(
            "O objeto fornecido não é uma planilha do openpyxl."
        )

    # Retorna a própria planilha validada
    return planilha


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
            Resultado produzido pela função.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Verifica se o valor informado é booleano
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
            Valor inteiro validado.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Impede que booleanos sejam interpretados como inteiros
    if isinstance(valor, bool) or not isinstance(valor, int):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser inteiro."
        )

    # Verifica se o valor respeita o limite mínimo
    if valor < minimo:
        # Interrompe a execução quando o valor é inválido
        raise ValueError(
            f"O parâmetro '{nome}' deve ser maior ou igual a {minimo}."
        )

    # Retorna o inteiro validado
    return valor


def _validar_numero(
    valor: int | float,
    nome: str,
    minimo: int | float | None = None,
    maximo: int | float | None = None
) -> int | float:
    """
    Valida um valor numérico e seus limites opcionais.

    Args:
        valor (int | float):
            Valor que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        minimo (int | float | None):
            Menor valor permitido para o parâmetro.
        maximo (int | float | None):
            Maior valor permitido para o parâmetro.

    Returns:
        int | float:
            Valor numérico validado.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Impede que booleanos sejam interpretados como números
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser numérico."
        )

    # Verifica o limite mínimo quando informado
    if minimo is not None and valor < minimo:
        # Interrompe a execução quando o valor é menor que o permitido
        raise DimensaoInvalidaError(
            f"O parâmetro '{nome}' deve ser maior ou igual a {minimo}."
        )

    # Verifica o limite máximo quando informado
    if maximo is not None and valor > maximo:
        # Interrompe a execução quando o valor ultrapassa o limite
        raise DimensaoInvalidaError(
            f"O parâmetro '{nome}' deve ser menor ou igual a {maximo}."
        )

    # Retorna o número validado
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
            Texto normalizado pela função.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Verifica se o valor informado é uma string
    if not isinstance(valor, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser uma string."
        )

    # Remove os espaços externos do texto
    texto = valor.strip()

    # Verifica se o texto possui conteúdo
    if not texto:
        # Interrompe a execução quando o texto está vazio
        raise ValueError(
            f"O parâmetro '{nome}' não pode estar vazio."
        )

    # Retorna o texto normalizado
    return texto


def _normalizar_referencia_intervalo(
    referencia: str
) -> str:
    """
    Valida e normaliza uma referência de célula ou intervalo.

    Args:
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.

    Returns:
        str:
            Texto normalizado pela função.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o texto da referência
    referencia_normalizada = _validar_texto(
        referencia,
        "referencia",
    ).upper()

    # Tenta interpretar os limites da referência
    try:
        minimo_coluna, minimo_linha, maximo_coluna, maximo_linha = (
            range_boundaries(referencia_normalizada)
        )
    except (TypeError, ValueError) as erro:
        # Converte a falha em uma exceção específica do módulo
        raise IntervaloFormatacaoInvalidoError(
            f"A referência de intervalo é inválida: '{referencia}'."
        ) from erro

    # Verifica se todos os limites foram identificados como números
    limites = (
        minimo_coluna,
        minimo_linha,
        maximo_coluna,
        maximo_linha
    )

    # Rejeita referências parciais, como somente uma coluna ou uma linha
    if any(limite is None for limite in limites):
        raise IntervaloFormatacaoInvalidoError(
            f"A referência de intervalo é inválida: '{referencia}'."
        )

    # Verifica se os limites calculados são positivos
    if min(limites) < 1:
        # Interrompe a execução quando o intervalo é inválido
        raise IntervaloFormatacaoInvalidoError(
            f"A referência de intervalo é inválida: '{referencia}'."
        )

    # Retorna a referência normalizada
    return referencia_normalizada


def _criar_lado_borda(
    estilo: str | None,
    cor: str
) -> Side:
    """
    Cria um lado de borda utilizando estilo e cor validados.

    Args:
        estilo (str | None):
            Estilo visual que será aplicado.
        cor (str):
            Código de cor RGB ou ARGB utilizado na formatação.

    Returns:
        Side:
            Objeto Side configurado para uso em bordas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Permite a criação de um lado sem estilo
    if estilo is None:
        return Side(style=None)

    # Valida o estilo textual da borda
    estilo_validado = _validar_texto(
        estilo,
        "estilo_borda"
    )

    # Verifica se o estilo pertence ao conjunto aceito
    if estilo_validado not in TIPOS_BORDA:
        # Interrompe a execução quando o estilo não é suportado
        raise EstiloInvalidoError(
            f"O estilo de borda não é suportado: '{estilo_validado}'."
        )

    # Cria e retorna o lado de borda configurado
    return Side(
        style=estilo_validado,
        color=normalizar_cor(cor)
    )


# ----------------------------------------------------------------------------
# FUNÇÕES DE CONSULTA
# ----------------------------------------------------------------------------

def normalizar_cor(cor: str) -> str:
    """
    Normaliza uma cor hexadecimal para o padrão ARGB do openpyxl.

    Args:
        cor (str):
            Código de cor RGB ou ARGB utilizado na formatação.

    Returns:
        str:
            Texto normalizado pela função.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida o texto da cor e remove o prefixo opcional
    cor_normalizada = _validar_texto(cor, "cor").lstrip("#").upper()

    # Verifica se todos os caracteres são hexadecimais
    if any(
        caractere not in "0123456789ABCDEF"
        for caractere in cor_normalizada
    ):
        # Interrompe a execução quando a cor possui caracteres inválidos
        raise CorInvalidaError(
            f"O código de cor é inválido: '{cor}'."
        )

    # Adiciona o canal alfa completo para cores RGB
    if len(cor_normalizada) == 6:
        cor_normalizada = f"FF{cor_normalizada}"

    # Verifica se o resultado possui oito caracteres
    if len(cor_normalizada) != 8:
        # Interrompe a execução quando o tamanho da cor é inválido
        raise CorInvalidaError(
            "A cor deve possuir seis caracteres RGB ou oito caracteres ARGB."
        )

    # Retorna o código ARGB normalizado
    return cor_normalizada


def obter_celulas_intervalo(
    planilha: Worksheet,
    referencia: str
) -> list[Cell]:
    """
    Retorna as células existentes em uma referência do Excel.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.

    Returns:
        list[Cell]:
            Lista linear com as células encontradas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Normaliza a referência informada
    referencia_normalizada = _normalizar_referencia_intervalo(referencia)

    # Obtém os limites numéricos do intervalo
    minimo_coluna, minimo_linha, maximo_coluna, maximo_linha = (
        range_boundaries(referencia_normalizada)
    )

    # Cria a lista que armazenará as células
    celulas: list[Cell] = []

    # Percorre todas as linhas do intervalo
    for indice_linha in range(minimo_linha, maximo_linha + 1):
        # Percorre todas as colunas do intervalo
        for indice_coluna in range(minimo_coluna, maximo_coluna + 1):
            # Adiciona a célula atual à lista
            celulas.append(
                planilha.cell(
                    row=indice_linha,
                    column=indice_coluna
                )
            )

    # Retorna a lista linear de células
    return celulas


# ----------------------------------------------------------------------------
# FUNÇÕES DE ESTILO
# ----------------------------------------------------------------------------

def aplicar_fonte(
    planilha: Worksheet,
    referencia: str,
    nome: str = "Calibri",
    tamanho: int | float = 11,
    negrito: bool = False,
    italico: bool = False,
    sublinhado: str | None = None,
    tachado: bool = False,
    cor: str = "FF000000"
) -> Worksheet:
    """
    Aplica configurações de fonte às células de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        tamanho (int | float):
            Tamanho da fonte que será aplicado.
        negrito (bool):
            Define se a fonte será exibida em negrito.
        italico (bool):
            Define se a fonte será exibida em itálico.
        sublinhado (str | None):
            Tipo opcional de sublinhado da fonte.
        tachado (bool):
            Define se a fonte será exibida com efeito tachado.
        cor (str):
            Código de cor RGB ou ARGB utilizado na formatação.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e os parâmetros da fonte
    _validar_planilha(planilha)
    nome_validado = _validar_texto(nome, "nome")
    tamanho_validado = _validar_numero(tamanho, "tamanho", minimo=1)

    # Valida as opções booleanas
    _validar_booleano(negrito, "negrito")
    _validar_booleano(italico, "italico")
    _validar_booleano(tachado, "tachado")

    # Valida o tipo opcional de sublinhado
    if sublinhado is not None and not isinstance(sublinhado, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'sublinhado' deve ser uma string ou None."
        )

    # Cria a fonte com as configurações recebidas
    fonte = Font(
        name=nome_validado,
        size=tamanho_validado,
        bold=negrito,
        italic=italico,
        underline=sublinhado,
        strike=tachado,
        color=normalizar_cor(cor)
    )

    # Percorre todas as células do intervalo
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Aplica uma cópia da fonte à célula atual
        celula.font = copy(fonte)

    # Retorna a planilha alterada
    return planilha


def aplicar_preenchimento(
    planilha: Worksheet,
    referencia: str,
    cor: str,
    tipo: str = "solid",
    cor_secundaria: str | None = None
) -> Worksheet:
    """
    Aplica preenchimento às células de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        cor (str):
            Código de cor RGB ou ARGB utilizado na formatação.
        tipo (str):
            Tipo de preenchimento utilizado.
        cor_secundaria (str | None):
            Cor secundária utilizada pelo preenchimento.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e os parâmetros textuais
    _validar_planilha(planilha)
    tipo_validado = _validar_texto(tipo, "tipo")

    # Normaliza as cores utilizadas pelo preenchimento
    cor_principal = normalizar_cor(cor)
    cor_final = (
        cor_principal
        if cor_secundaria is None
        else normalizar_cor(cor_secundaria)
    )

    # Cria o preenchimento configurado
    preenchimento = PatternFill(
        fill_type=tipo_validado,
        fgColor=cor_principal,
        bgColor=cor_final
    )

    # Percorre as células selecionadas
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Aplica uma cópia do preenchimento
        celula.fill = copy(preenchimento)

    # Retorna a planilha alterada
    return planilha


def aplicar_alinhamento(
    planilha: Worksheet,
    referencia: str,
    horizontal: str = "general",
    vertical: str = "bottom",
    quebrar_texto: bool = False,
    reduzir_para_caber: bool = False,
    recuo: int = 0,
    rotacao: int = 0
) -> Worksheet:
    """
    Aplica alinhamento horizontal e vertical a um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        horizontal (str):
            Alinhamento horizontal aplicado às células.
        vertical (str):
            Alinhamento vertical aplicado às células.
        quebrar_texto (bool):
            Define se o texto poderá ser quebrado em várias linhas.
        reduzir_para_caber (bool):
            Define se o conteúdo será reduzido para caber na célula.
        recuo (int):
            Quantidade de recuo aplicada ao conteúdo.
        rotacao (int):
            Ângulo de rotação aplicado ao texto.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e normaliza os alinhamentos
    _validar_planilha(planilha)
    horizontal_validado = _validar_texto(horizontal, "horizontal")
    vertical_validado = _validar_texto(vertical, "vertical")

    # Verifica se o alinhamento horizontal é suportado
    if horizontal_validado not in ALINHAMENTOS_HORIZONTAIS:
        # Interrompe a execução quando o alinhamento é inválido
        raise EstiloInvalidoError(
            f"Alinhamento horizontal inválido: '{horizontal_validado}'."
        )

    # Verifica se o alinhamento vertical é suportado
    if vertical_validado not in ALINHAMENTOS_VERTICAIS:
        # Interrompe a execução quando o alinhamento é inválido
        raise EstiloInvalidoError(
            f"Alinhamento vertical inválido: '{vertical_validado}'."
        )

    # Valida as opções booleanas e numéricas
    _validar_booleano(quebrar_texto, "quebrar_texto")
    _validar_booleano(reduzir_para_caber, "reduzir_para_caber")
    recuo_validado = _validar_inteiro(recuo, "recuo", minimo=0)
    rotacao_validada = _validar_inteiro(rotacao, "rotacao", minimo=0)

    # Verifica o limite aceito para rotação de texto
    if rotacao_validada > 180:
        # Interrompe a execução quando a rotação ultrapassa o limite
        raise EstiloInvalidoError(
            "A rotação do texto deve estar entre 0 e 180 graus."
        )

    # Cria o objeto de alinhamento
    alinhamento = Alignment(
        horizontal=horizontal_validado,
        vertical=vertical_validado,
        wrap_text=quebrar_texto,
        shrink_to_fit=reduzir_para_caber,
        indent=recuo_validado,
        text_rotation=rotacao_validada
    )

    # Percorre as células do intervalo
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Aplica uma cópia do alinhamento
        celula.alignment = copy(alinhamento)

    # Retorna a planilha alterada
    return planilha


def aplicar_bordas(
    planilha: Worksheet,
    referencia: str,
    estilo: str = "thin",
    cor: str = COR_BORDA_PADRAO,
    esquerda: bool = True,
    direita: bool = True,
    superior: bool = True,
    inferior: bool = True
) -> Worksheet:
    """
    Aplica bordas às células de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        estilo (str):
            Estilo visual que será aplicado.
        cor (str):
            Código de cor RGB ou ARGB utilizado na formatação.
        esquerda (bool):
            Define se a borda esquerda será aplicada.
        direita (bool):
            Define se a borda direita será aplicada.
        superior (bool):
            Define se a borda superior será aplicada.
        inferior (bool):
            Define se a borda inferior será aplicada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Valida as opções de cada lado da borda
    for nome, valor in {
        "esquerda": esquerda,
        "direita": direita,
        "superior": superior,
        "inferior": inferior,
    }.items():
        _validar_booleano(valor, nome)

    # Cria o lado de borda ativo
    lado = _criar_lado_borda(estilo, cor)

    # Cria o lado sem borda
    sem_borda = Side(style=None)

    # Cria o objeto Border com os lados selecionados
    borda = Border(
        left=lado if esquerda else sem_borda,
        right=lado if direita else sem_borda,
        top=lado if superior else sem_borda,
        bottom=lado if inferior else sem_borda
    )

    # Percorre as células selecionadas
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Aplica uma cópia da borda
        celula.border = copy(borda)

    # Retorna a planilha alterada
    return planilha


def aplicar_formato_numero(
    planilha: Worksheet,
    referencia: str,
    formato: str = FORMATO_NUMERO_PADRAO
) -> Worksheet:
    """
    Aplica um código de formato numérico a um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        formato (str):
            Código de formato numérico aplicado às células.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e o formato informado
    _validar_planilha(planilha)
    formato_validado = _validar_texto(formato, "formato")

    # Percorre as células selecionadas
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Aplica o mesmo código de formato numérico
        celula.number_format = formato_validado

    # Retorna a planilha alterada
    return planilha


def aplicar_formato_data(
    planilha: Worksheet,
    referencia: str,
    formato: str = FORMATO_DATA_PADRAO
) -> Worksheet:
    """
    Aplica um formato de data às células selecionadas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        formato (str):
            Código de formato numérico aplicado às células.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Reutiliza a função genérica de formato numérico
    return aplicar_formato_numero(
        planilha,
        referencia,
        formato=formato
    )


def aplicar_formato_moeda(
    planilha: Worksheet,
    referencia: str,
    formato: str = FORMATO_MOEDA_PADRAO
) -> Worksheet:
    """
    Aplica um formato monetário às células selecionadas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        formato (str):
            Código de formato numérico aplicado às células.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Reutiliza a função genérica de formato numérico
    return aplicar_formato_numero(
        planilha,
        referencia,
        formato=formato
    )


def aplicar_formato_percentual(
    planilha: Worksheet,
    referencia: str,
    formato: str = FORMATO_PERCENTUAL_PADRAO
) -> Worksheet:
    """
    Aplica um formato percentual às células selecionadas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        formato (str):
            Código de formato numérico aplicado às células.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Reutiliza a função genérica de formato numérico
    return aplicar_formato_numero(
        planilha,
        referencia,
        formato=formato
    )


def formatar_cabecalho(
    planilha: Worksheet,
    referencia: str,
    cor_fundo: str = COR_CABECALHO_PADRAO,
    cor_texto: str = COR_TEXTO_CABECALHO_PADRAO,
    tamanho_fonte: int | float = 11,
    alinhamento_horizontal: str = "center",
    alinhamento_vertical: str = "center",
    aplicar_borda_inferior: bool = True
) -> Worksheet:
    """
    Aplica um estilo completo de cabeçalho ao intervalo informado.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        cor_fundo (str):
            Cor utilizada no preenchimento do cabeçalho.
        cor_texto (str):
            Cor utilizada no texto do cabeçalho.
        tamanho_fonte (int | float):
            Tamanho da fonte utilizada no cabeçalho.
        alinhamento_horizontal (str):
            Alinhamento horizontal do cabeçalho.
        alinhamento_vertical (str):
            Alinhamento vertical do cabeçalho.
        aplicar_borda_inferior (bool):
            Define se a borda inferior será aplicada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a opção de borda inferior
    _validar_booleano(
        aplicar_borda_inferior,
        "aplicar_borda_inferior"
    )

    # Aplica a fonte do cabeçalho
    aplicar_fonte(
        planilha,
        referencia,
        tamanho=tamanho_fonte,
        negrito=True,
        cor=cor_texto
    )

    # Aplica o preenchimento do cabeçalho
    aplicar_preenchimento(
        planilha,
        referencia,
        cor=cor_fundo,
        tipo="solid"
    )

    # Aplica o alinhamento e a quebra automática de texto
    aplicar_alinhamento(
        planilha,
        referencia,
        horizontal=alinhamento_horizontal,
        vertical=alinhamento_vertical,
        quebrar_texto=True
    )

    # Aplica somente a borda inferior quando solicitado
    if aplicar_borda_inferior:
        aplicar_bordas(
            planilha,
            referencia,
            estilo="thin",
            cor="FF000000",
            esquerda=False,
            direita=False,
            superior=False,
            inferior=True
        )

    # Retorna a planilha alterada
    return planilha


def formatar_intervalo(
    planilha: Worksheet,
    referencia: str,
    fonte: dict[str, Any] | None = None,
    preenchimento: dict[str, Any] | None = None,
    alinhamento: dict[str, Any] | None = None,
    bordas: dict[str, Any] | None = None,
    formato_numero: str | None = None,
    bloqueado: bool | None = None
) -> Worksheet:
    """
    Aplica várias configurações de estilo em uma única operação.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        fonte (dict[str, Any] | None):
            Dicionário com as configurações encaminhadas à função de fonte.
        preenchimento (dict[str, Any] | None):
            Dicionário com as configurações encaminhadas à função de preenchimento.
        alinhamento (dict[str, Any] | None):
            Dicionário com as configurações encaminhadas à função de alinhamento.
        bordas (dict[str, Any] | None):
            Dicionário com as configurações encaminhadas à função de bordas.
        formato_numero (str | None):
            Código opcional de formato numérico.
        bloqueado (bool | None):
            Define opcionalmente se as células ficarão bloqueadas.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a referência antes das operações
    _validar_planilha(planilha)
    _normalizar_referencia_intervalo(referencia)

    # Valida os dicionários opcionais
    for nome, configuracao in {
        "fonte": fonte,
        "preenchimento": preenchimento,
        "alinhamento": alinhamento,
        "bordas": bordas,
    }.items():
        # Verifica se cada configuração é um dicionário ou None
        if configuracao is not None and not isinstance(configuracao, dict):
            # Interrompe a execução quando o tipo é inválido
            raise TypeError(
                f"O parâmetro '{nome}' deve ser um dicionário ou None."
            )

    # Aplica as configurações de fonte quando informadas
    if fonte is not None:
        aplicar_fonte(planilha, referencia, **fonte)

    # Aplica as configurações de preenchimento quando informadas
    if preenchimento is not None:
        aplicar_preenchimento(planilha, referencia, **preenchimento)

    # Aplica as configurações de alinhamento quando informadas
    if alinhamento is not None:
        aplicar_alinhamento(planilha, referencia, **alinhamento)

    # Aplica as configurações de borda quando informadas
    if bordas is not None:
        aplicar_bordas(planilha, referencia, **bordas)

    # Aplica o formato numérico quando informado
    if formato_numero is not None:
        aplicar_formato_numero(
            planilha,
            referencia,
            formato=formato_numero
        )

    # Aplica a proteção quando informada
    if bloqueado is not None:
        _validar_booleano(bloqueado, "bloqueado")

        # Percorre as células selecionadas
        for celula in obter_celulas_intervalo(planilha, referencia):
            # Preserva as demais opções da proteção atual
            celula.protection = Protection(
                locked=bloqueado,
                hidden=celula.protection.hidden
            )

    # Retorna a planilha alterada
    return planilha


def copiar_formatacao(
    planilha: Worksheet,
    origem: str,
    destino: str
) -> Worksheet:
    """
    Copia a formatação de uma célula para outra célula.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        origem (str):
            Referência da célula de origem da formatação.
        destino (str):
            Referência da célula de destino da formatação.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém as células das referências informadas
    celulas_origem = obter_celulas_intervalo(planilha, origem)
    celulas_destino = obter_celulas_intervalo(planilha, destino)

    # Exige referências contendo apenas uma célula
    if len(celulas_origem) != 1 or len(celulas_destino) != 1:
        # Interrompe a execução quando alguma referência é um intervalo
        raise IntervaloFormatacaoInvalidoError(
            "As referências de origem e destino devem conter uma única célula."
        )

    # Obtém as células individuais
    celula_origem = celulas_origem[0]
    celula_destino = celulas_destino[0]

    # Copia todos os principais componentes de estilo
    celula_destino.font = copy(celula_origem.font)
    celula_destino.fill = copy(celula_origem.fill)
    celula_destino.border = copy(celula_origem.border)
    celula_destino.alignment = copy(celula_origem.alignment)
    celula_destino.number_format = celula_origem.number_format
    celula_destino.protection = copy(celula_origem.protection)

    # Retorna a planilha alterada
    return planilha


def limpar_formatacao(
    planilha: Worksheet,
    referencia: str,
    manter_formato_numero: bool = False
) -> Worksheet:
    """
    Restaura a formatação padrão das células de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        manter_formato_numero (bool):
            Define se o formato numérico atual será preservado.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a opção de preservação
    _validar_planilha(planilha)
    _validar_booleano(
        manter_formato_numero,
        "manter_formato_numero"
    )

    # Percorre as células selecionadas
    for celula in obter_celulas_intervalo(planilha, referencia):
        # Armazena o formato numérico quando precisa ser preservado
        formato_anterior = celula.number_format

        # Cria uma célula temporária com o estilo padrão
        celula_padrao = Cell(
            planilha,
            row=celula.row,
            column=celula.column
        )

        # Restaura os componentes visuais padrão
        celula.font = copy(celula_padrao.font)
        celula.fill = copy(celula_padrao.fill)
        celula.border = copy(celula_padrao.border)
        celula.alignment = copy(celula_padrao.alignment)
        celula.protection = copy(celula_padrao.protection)
        celula.number_format = (
            formato_anterior
            if manter_formato_numero
            else celula_padrao.number_format
        )

    # Retorna a planilha alterada
    return planilha


# ----------------------------------------------------------------------------
# FUNÇÕES DE DIMENSÕES
# ----------------------------------------------------------------------------

def definir_largura_coluna(
    planilha: Worksheet,
    coluna: str | int,
    largura: int | float
) -> Worksheet:
    """
    Define a largura de uma coluna da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        coluna (str | int):
            Letra ou índice numérico da coluna.
        largura (int | float):
            Largura que será aplicada à coluna.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Normaliza a referência da coluna
    if isinstance(coluna, int) and not isinstance(coluna, bool):
        # Valida o índice e converte para letra
        indice = _validar_inteiro(coluna, "coluna", minimo=1)
        letra = get_column_letter(indice)
    elif isinstance(coluna, str):
        # Valida e normaliza a letra da coluna
        letra = _validar_texto(coluna, "coluna").upper()

        # Verifica se a referência possui somente letras
        if not letra.isalpha():
            raise ValueError(
                "A referência da coluna deve possuir somente letras."
            )
    else:
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'coluna' deve ser uma string ou um inteiro."
        )

    # Valida a largura dentro dos limites do Excel
    largura_validada = _validar_numero(
        largura,
        "largura",
        minimo=0,
        maximo=LARGURA_MAXIMA_COLUNA
    )

    # Define a largura da coluna
    planilha.column_dimensions[letra].width = largura_validada

    # Retorna a planilha alterada
    return planilha


def definir_altura_linha(
    planilha: Worksheet,
    linha: int,
    altura: int | float
) -> Worksheet:
    """
    Define a altura de uma linha da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        linha (int):
            Índice numérico da linha.
        altura (int | float):
            Altura que será aplicada à linha.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e o índice da linha
    _validar_planilha(planilha)
    linha_validada = _validar_inteiro(linha, "linha", minimo=1)

    # Valida a altura dentro dos limites do Excel
    altura_validada = _validar_numero(
        altura,
        "altura",
        minimo=0,
        maximo=ALTURA_MAXIMA_LINHA
    )

    # Define a altura da linha
    planilha.row_dimensions[linha_validada].height = altura_validada

    # Retorna a planilha alterada
    return planilha


def ajustar_larguras_colunas(
    planilha: Worksheet,
    colunas: Iterable[str | int] | None = None,
    linha_inicial: int = 1,
    linha_final: int | None = None,
    margem: int | float = 2,
    largura_minima: int | float = 1,
    largura_maxima: int | float = 80
) -> dict[str, float]:
    """
    Ajusta automaticamente a largura de várias colunas.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        colunas (Iterable[str | int] | None):
            Coleção opcional de colunas que serão ajustadas.
        linha_inicial (int):
            Primeira linha incluída na operação.
        linha_final (int | None):
            Última linha incluída na operação.
        margem (int | float):
            Margem adicional utilizada no cálculo da largura.
        largura_minima (int | float):
            Menor largura permitida.
        largura_maxima (int | float):
            Maior largura permitida.

    Returns:
        dict[str, float]:
            Dicionário com as letras e larguras aplicadas.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e os limites numéricos
    _validar_planilha(planilha)
    inicio = _validar_inteiro(
        linha_inicial,
        "linha_inicial",
        minimo=1
    )
    fim = (
        planilha.max_row
        if linha_final is None
        else _validar_inteiro(linha_final, "linha_final", minimo=1)
    )

    # Verifica se o intervalo de linhas está em ordem crescente
    if fim < inicio:
        raise ValueError(
            "A linha final não pode ser menor que a linha inicial."
        )

    # Valida os parâmetros utilizados no cálculo
    margem_validada = _validar_numero(
        margem,
        "margem",
        minimo=0
    )
    minima_validada = _validar_numero(
        largura_minima,
        "largura_minima",
        minimo=0,
        maximo=LARGURA_MAXIMA_COLUNA
    )
    maxima_validada = _validar_numero(
        largura_maxima,
        "largura_maxima",
        minimo=0,
        maximo=LARGURA_MAXIMA_COLUNA
    )

    # Verifica se os limites de largura estão consistentes
    if maxima_validada < minima_validada:
        raise DimensaoInvalidaError(
            "A largura máxima não pode ser menor que a largura mínima."
        )

    # Define as colunas analisadas quando nenhuma seleção foi informada
    if colunas is None:
        referencias: list[str | int] = list(
            range(1, planilha.max_column + 1)
        )
    else:
        # Rejeita strings isoladas para evitar uma coluna por caractere
        if isinstance(colunas, (str, bytes)) or not isinstance(
            colunas,
            Iterable,
        ):
            raise TypeError(
                "O parâmetro 'colunas' deve ser um iterável não textual ou None."
            )

        # Converte o iterável para lista
        referencias = list(colunas)

    # Cria o resultado das larguras aplicadas
    larguras: dict[str, float] = {}

    # Percorre as referências selecionadas
    for referencia in referencias:
        # Normaliza a referência para índice e letra
        if isinstance(referencia, int) and not isinstance(referencia, bool):
            indice_coluna = _validar_inteiro(
                referencia,
                "coluna",
                minimo=1
            )
            letra_coluna = get_column_letter(indice_coluna)
        elif isinstance(referencia, str):
            letra_coluna = _validar_texto(
                referencia,
                "coluna",
            ).upper()

            # Converte a letra em índice utilizando o intervalo de uma célula
            try:
                indice_coluna = range_boundaries(
                    f"{letra_coluna}1"
                )[0]
            except (TypeError, ValueError) as erro:
                raise ValueError(
                    f"A referência da coluna é inválida: '{referencia}'."
                ) from erro
        else:
            raise TypeError(
                "Cada referência de coluna deve ser string ou inteiro."
            )

        # Inicia o maior comprimento encontrado
        maior_comprimento = 0

        # Percorre as linhas selecionadas
        for indice_linha in range(inicio, fim + 1):
            # Obtém o valor da célula atual
            valor = planilha.cell(
                row=indice_linha,
                column=indice_coluna,
            ).value

            # Ignora células vazias
            if valor is None:
                continue

            # Calcula o maior trecho do conteúdo
            comprimento = max(
                len(parte)
                for parte in str(valor).splitlines() or [""]
            )

            # Atualiza o maior comprimento encontrado
            maior_comprimento = max(
                maior_comprimento,
                comprimento
            )

        # Calcula a largura respeitando margem e limites
        largura = max(
            minima_validada,
            min(
                maior_comprimento + margem_validada,
                maxima_validada,
            )
        )

        # Aplica a largura calculada
        planilha.column_dimensions[letra_coluna].width = largura

        # Armazena a largura no resultado
        larguras[letra_coluna] = float(largura)

    # Retorna as larguras aplicadas
    return larguras


# ----------------------------------------------------------------------------
# FUNÇÕES DE VISUALIZAÇÃO
# ----------------------------------------------------------------------------

def congelar_paineis(
    planilha: Worksheet,
    referencia: str = "A2"
) -> Worksheet:
    """
    Congela linhas e colunas acima e à esquerda da referência.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e exige uma referência de célula única
    _validar_planilha(planilha)
    celulas = obter_celulas_intervalo(planilha, referencia)

    # Verifica se a referência contém somente uma célula
    if len(celulas) != 1:
        raise IntervaloFormatacaoInvalidoError(
            "A referência de congelamento deve conter uma única célula."
        )

    # Define o ponto de congelamento
    planilha.freeze_panes = celulas[0].coordinate

    # Retorna a planilha alterada
    return planilha


def descongelar_paineis(planilha: Worksheet) -> Worksheet:
    """
    Remove o congelamento de painéis da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Remove a referência de congelamento
    planilha.freeze_panes = None

    # Retorna a planilha alterada
    return planilha


def mesclar_celulas(
    planilha: Worksheet,
    referencia: str
) -> Worksheet:
    """
    Mescla as células do intervalo informado.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a referência
    _validar_planilha(planilha)
    referencia_normalizada = _normalizar_referencia_intervalo(referencia)

    # Exige mais de uma célula para realizar a mesclagem
    if len(obter_celulas_intervalo(planilha, referencia_normalizada)) < 2:
        raise IntervaloFormatacaoInvalidoError(
            "A mesclagem exige um intervalo com mais de uma célula."
        )

    # Mescla o intervalo selecionado
    planilha.merge_cells(referencia_normalizada)

    # Retorna a planilha alterada
    return planilha


def desmesclar_celulas(
    planilha: Worksheet,
    referencia: str
) -> Worksheet:
    """
    Desfaz a mesclagem do intervalo informado.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a referência
    _validar_planilha(planilha)
    referencia_normalizada = _normalizar_referencia_intervalo(referencia)

    # Verifica se o intervalo está entre as mesclagens existentes
    intervalos_mesclados = {
        str(intervalo)
        for intervalo in planilha.merged_cells.ranges
    }

    # Interrompe a execução quando o intervalo não está mesclado
    if referencia_normalizada not in intervalos_mesclados:
        raise IntervaloFormatacaoInvalidoError(
            f"O intervalo não está mesclado: '{referencia_normalizada}'."
        )

    # Desfaz a mesclagem
    planilha.unmerge_cells(referencia_normalizada)

    # Retorna a planilha alterada
    return planilha


def ocultar_linhas_grade(planilha: Worksheet) -> Worksheet:
    """
    Oculta as linhas de grade da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Desativa a exibição das linhas de grade
    planilha.sheet_view.showGridLines = False

    # Retorna a planilha alterada
    return planilha


def exibir_linhas_grade(planilha: Worksheet) -> Worksheet:
    """
    Exibe novamente as linhas de grade da planilha.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Ativa a exibição das linhas de grade
    planilha.sheet_view.showGridLines = True

    # Retorna a planilha alterada
    return planilha


def aplicar_estilo_zebrado(
    planilha: Worksheet,
    referencia: str,
    cor_linhas_pares: str = COR_LINHA_ALTERNADA_PADRAO,
    cor_linhas_impares: str | None = None,
    primeira_linha_dados: int | None = None
) -> Worksheet:
    """
    Aplica cores alternadas às linhas de um intervalo.

    Args:
        planilha (Worksheet):
            Planilha do openpyxl que será validada, consultada ou alterada.
        referencia (str):
            Referência da célula ou intervalo no padrão do Excel.
        cor_linhas_pares (str):
            Cor aplicada às linhas alternadas pares.
        cor_linhas_impares (str | None):
            Cor opcional aplicada às linhas alternadas ímpares.
        primeira_linha_dados (int | None):
            Primeira linha considerada no padrão zebrado.

    Returns:
        Worksheet:
            A própria planilha alterada ou validada.

    Raises:
        TypeError:
            Caso algum parâmetro possua tipo incompatível.
        ValueError:
            Caso algum valor informado seja inválido.
    """

    # Valida a planilha e a referência
    _validar_planilha(planilha)
    referencia_normalizada = _normalizar_referencia_intervalo(referencia)

    # Obtém os limites do intervalo
    minimo_coluna, minimo_linha, maximo_coluna, maximo_linha = (
        range_boundaries(referencia_normalizada)
    )

    # Define a primeira linha considerada no padrão alternado
    linha_base = (
        minimo_linha
        if primeira_linha_dados is None
        else _validar_inteiro(
            primeira_linha_dados,
            "primeira_linha_dados",
            minimo=minimo_linha
        )
    )

    # Verifica se a linha base pertence ao intervalo
    if linha_base > maximo_linha:
        raise ValueError(
            "A primeira linha de dados deve pertencer ao intervalo informado."
        )

    # Cria o preenchimento das linhas pares
    preenchimento_par = PatternFill(
        fill_type="solid",
        fgColor=normalizar_cor(cor_linhas_pares)
    )

    # Cria o preenchimento das linhas ímpares quando uma cor foi informada
    preenchimento_impar = (
        None
        if cor_linhas_impares is None
        else PatternFill(
            fill_type="solid",
            fgColor=normalizar_cor(cor_linhas_impares)
        )
    )

    # Percorre as linhas de dados do intervalo
    for indice_linha in range(linha_base, maximo_linha + 1):
        # Identifica se a posição relativa da linha é par
        linha_par = (indice_linha - linha_base) % 2 == 1

        # Seleciona o preenchimento adequado
        preenchimento = (
            preenchimento_par
            if linha_par
            else preenchimento_impar
        )

        # Percorre todas as colunas da linha atual
        for indice_coluna in range(minimo_coluna, maximo_coluna + 1):
            # Obtém a célula atual
            celula = planilha.cell(
                row=indice_linha,
                column=indice_coluna
            )

            # Aplica a cor quando existe preenchimento configurado
            if preenchimento is not None:
                celula.fill = copy(preenchimento)

    # Retorna a planilha alterada
    return planilha


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

# Define explicitamente os objetos públicos disponibilizados pelo módulo
__all__ = [
    "COR_CABECALHO_PADRAO",
    "COR_TEXTO_CABECALHO_PADRAO",
    "COR_LINHA_ALTERNADA_PADRAO",
    "COR_BORDA_PADRAO",
    "FORMATO_DATA_PADRAO",
    "FORMATO_MOEDA_PADRAO",
    "FORMATO_PERCENTUAL_PADRAO",
    "FORMATO_NUMERO_PADRAO",
    "LARGURA_MAXIMA_COLUNA",
    "ALTURA_MAXIMA_LINHA",
    "TIPOS_BORDA",
    "ALINHAMENTOS_HORIZONTAIS",
    "ALINHAMENTOS_VERTICAIS",
    "ExcelFormatacaoError",
    "PlanilhaFormatacaoInvalidaError",
    "IntervaloFormatacaoInvalidoError",
    "CorInvalidaError",
    "EstiloInvalidoError",
    "DimensaoInvalidaError",
    "normalizar_cor",
    "obter_celulas_intervalo",
    "aplicar_fonte",
    "aplicar_preenchimento",
    "aplicar_alinhamento",
    "aplicar_bordas",
    "aplicar_formato_numero",
    "aplicar_formato_data",
    "aplicar_formato_moeda",
    "aplicar_formato_percentual",
    "formatar_cabecalho",
    "formatar_intervalo",
    "copiar_formatacao",
    "limpar_formatacao",
    "definir_largura_coluna",
    "definir_altura_linha",
    "ajustar_larguras_colunas",
    "congelar_paineis",
    "descongelar_paineis",
    "mesclar_celulas",
    "desmesclar_celulas",
    "ocultar_linhas_grade",
    "exibir_linhas_grade",
    "aplicar_estilo_zebrado",
]
