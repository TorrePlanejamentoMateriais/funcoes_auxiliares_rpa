"""
============================================================
Módulo: EXCEL / colunas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para consultar, inserir, remover,
    copiar, mover, ocultar, exibir e dimensionar colunas de planilhas
    Excel utilizando openpyxl.
Funções disponíveis:
    - converter_coluna_para_numero()
    - converter_numero_para_coluna()
    - normalizar_coluna()
    - obter_indice_coluna()
    - obter_letra_coluna()
    - obter_ultima_coluna()
    - obter_valores_coluna()
    - definir_valores_coluna()
    - inserir_colunas()
    - remover_colunas()
    - copiar_coluna()
    - mover_coluna()
    - limpar_coluna()
    - ocultar_coluna()
    - exibir_coluna()
    - definir_largura_coluna()
    - obter_largura_coluna()
    - ajustar_largura_coluna()
    - ajustar_larguras_colunas()
    - listar_colunas_preenchidas()
    - localizar_coluna_por_cabecalho()
    - renomear_cabecalho_coluna()
Dependências:
    - openpyxl
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão das funções de conversão entre letras e números.
        - Inclusão das operações de inserção, remoção, cópia e movimento.
        - Inclusão das funções de largura, ocultação e cabeçalhos.
        - Inclusão de validações para planilhas, linhas e colunas.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa copy para duplicar estilos sem compartilhar objetos mutáveis
from copy import copy

# Importa tipos auxiliares utilizados nas anotações das funções
from typing import Any, Iterable

# Importa utilitários para conversão entre índices e letras de colunas
from openpyxl.utils import get_column_letter, column_index_from_string

# Importa a classe Worksheet para validar as planilhas recebidas
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# EXCEÇÕES
# ----------------------------------------------------------------------------

class ExcelColunasError(Exception):
    """Erro base para operações relacionadas a colunas do Excel."""


class PlanilhaInvalidaError(ExcelColunasError):
    """A planilha informada não é um objeto Worksheet válido."""


class ColunaInvalidaError(ExcelColunasError, ValueError):
    """A referência de coluna informada é inválida."""


class ColunaNaoEncontradaError(ExcelColunasError, KeyError):
    """A coluna ou o cabeçalho solicitado não foi encontrado."""


class OperacaoColunaError(ExcelColunasError):
    """Uma operação de alteração da coluna não pôde ser concluída."""


# ----------------------------------------------------------------------------
# FUNÇÕES INTERNAS DE VALIDAÇÃO
# ----------------------------------------------------------------------------

def _validar_planilha(planilha: Worksheet) -> Worksheet:
    """
    Valida se o objeto informado é uma planilha do openpyxl.

    Args:
        planilha (Worksheet):
            Planilha que será validada.

    Returns:
        Worksheet:
            A própria planilha validada.

    Raises:
        PlanilhaInvalidaError:
            Caso o objeto não seja uma planilha do openpyxl.
    """

    # Verifica se o objeto recebido é uma planilha válida
    if not isinstance(planilha, Worksheet):
        # Interrompe a execução quando o objeto não é uma Worksheet
        raise PlanilhaInvalidaError(
            "O objeto fornecido não é uma planilha do openpyxl."
        )

    # Retorna a planilha validada
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
            O próprio valor validado.

    Raises:
        TypeError:
            Caso o valor não seja booleano.
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
    minimo: int = 1,
) -> int:
    """
    Valida um número inteiro e seu limite mínimo.

    Args:
        valor (int):
            Número que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        minimo (int):
            Menor valor permitido.

    Returns:
        int:
            O número inteiro validado.

    Raises:
        TypeError:
            Caso o valor não seja inteiro ou seja booleano.
        ValueError:
            Caso o valor seja menor que o limite mínimo.
    """

    # Verifica se o valor é um inteiro e impede booleanos
    if isinstance(valor, bool) or not isinstance(valor, int):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser inteiro."
        )

    # Verifica se o valor respeita o limite mínimo
    if valor < minimo:
        # Interrompe a execução quando o valor é menor que o permitido
        raise ValueError(
            f"O parâmetro '{nome}' deve ser maior ou igual a {minimo}."
        )

    # Retorna o inteiro validado
    return valor


def _validar_texto(
    valor: str,
    nome: str,
    permitir_vazio: bool = False,
) -> str:
    """
    Valida e normaliza um parâmetro textual.

    Args:
        valor (str):
            Texto que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        permitir_vazio (bool):
            Define se textos vazios podem ser aceitos.

    Returns:
        str:
            Texto sem espaços externos.

    Raises:
        TypeError:
            Caso o valor não seja uma string.
        ValueError:
            Caso o texto esteja vazio e isso não seja permitido.
    """

    # Valida a opção que controla textos vazios
    _validar_booleano(permitir_vazio, "permitir_vazio")

    # Verifica se o valor informado é uma string
    if not isinstance(valor, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser uma string."
        )

    # Remove os espaços externos do texto
    texto = valor.strip()

    # Verifica se o texto ficou vazio
    if not texto and not permitir_vazio:
        # Interrompe a execução quando um texto obrigatório está vazio
        raise ValueError(
            f"O parâmetro '{nome}' não pode estar vazio."
        )

    # Retorna o texto normalizado
    return texto


def _copiar_estilo_celula(origem: Any, destino: Any) -> None:
    """
    Copia o estilo de uma célula para outra célula.

    Args:
        origem (Cell):
            Célula que contém o estilo original.
        destino (Cell):
            Célula que receberá o estilo copiado.
    """

    # Copia a fonte da célula original
    destino.font = copy(origem.font)

    # Copia o preenchimento da célula original
    destino.fill = copy(origem.fill)

    # Copia as bordas da célula original
    destino.border = copy(origem.border)

    # Copia o alinhamento da célula original
    destino.alignment = copy(origem.alignment)

    # Copia o formato numérico da célula original
    destino.number_format = origem.number_format

    # Copia a proteção da célula original
    destino.protection = copy(origem.protection)


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE CONVERSÃO E CONSULTA
# ----------------------------------------------------------------------------

def converter_coluna_para_numero(coluna: str) -> int:
    """
    Converte a letra de uma coluna do Excel para seu índice numérico.

    Args:
        coluna (str):
            Letra da coluna que será convertida, como A, B ou AA.

    Returns:
        int:
            Índice numérico da coluna, iniciado em 1.

    Raises:
        TypeError:
            Caso a coluna não seja uma string.
        ColunaInvalidaError:
            Caso a referência não represente uma coluna válida.
    """

    # Valida e converte a referência para letras maiúsculas
    coluna_normalizada = _validar_texto(
        coluna,
        "coluna",
    ).upper()

    # Verifica se a referência possui somente letras
    if not coluna_normalizada.isalpha():
        # Interrompe a execução quando a referência contém outros caracteres
        raise ColunaInvalidaError(
            f"A referência de coluna é inválida: '{coluna}'."
        )

    # Executa a conversão através do utilitário do openpyxl
    try:
        indice = column_index_from_string(coluna_normalizada)
    except ValueError as erro:
        # Converte a falha original em uma exceção específica do módulo
        raise ColunaInvalidaError(
            f"A referência de coluna é inválida: '{coluna}'."
        ) from erro

    # Retorna o índice numérico da coluna
    return indice


def converter_numero_para_coluna(numero: int) -> str:
    """
    Converte um índice numérico para a letra correspondente do Excel.

    Args:
        numero (int):
            Índice numérico da coluna, iniciado em 1.

    Returns:
        str:
            Letra correspondente da coluna.

    Raises:
        TypeError:
            Caso o número não seja inteiro.
        ValueError:
            Caso o número seja menor que 1.
        ColunaInvalidaError:
            Caso o índice ultrapasse o limite suportado pelo Excel.
    """

    # Valida o índice numérico recebido
    indice = _validar_inteiro(
        numero,
        "numero",
        minimo=1,
    )

    # Executa a conversão através do utilitário do openpyxl
    try:
        letra = get_column_letter(indice)
    except ValueError as erro:
        # Converte a falha original em uma exceção específica do módulo
        raise ColunaInvalidaError(
            f"O índice de coluna é inválido: {numero}."
        ) from erro

    # Retorna a letra correspondente da coluna
    return letra


def normalizar_coluna(coluna: str | int) -> tuple[int, str]:
    """
    Normaliza uma referência de coluna em índice e letra.

    Args:
        coluna (str | int):
            Letra ou índice numérico da coluna.

    Returns:
        tuple[int, str]:
            Tupla contendo o índice numérico e a letra da coluna.

    Raises:
        TypeError:
            Caso a referência não seja string ou inteiro.
    """

    # Verifica se a referência foi informada como número
    if isinstance(coluna, int) and not isinstance(coluna, bool):
        # Valida o índice e obtém sua letra correspondente
        indice = _validar_inteiro(coluna, "coluna", minimo=1)
        letra = converter_numero_para_coluna(indice)

        # Retorna as duas representações da coluna
        return indice, letra

    # Verifica se a referência foi informada como texto
    if isinstance(coluna, str):
        # Converte a letra para índice e normaliza para maiúsculas
        letra = _validar_texto(coluna, "coluna").upper()
        indice = converter_coluna_para_numero(letra)

        # Retorna as duas representações da coluna
        return indice, letra

    # Interrompe a execução quando o tipo é inválido
    raise TypeError(
        "O parâmetro 'coluna' deve ser uma string ou um inteiro."
    )


def obter_indice_coluna(coluna: str | int) -> int:
    """Retorna o índice numérico de uma referência de coluna."""

    # Normaliza a referência e seleciona o índice
    indice, _ = normalizar_coluna(coluna)

    # Retorna o índice numérico
    return indice


def obter_letra_coluna(coluna: str | int) -> str:
    """Retorna a letra de uma referência de coluna."""

    # Normaliza a referência e seleciona a letra
    _, letra = normalizar_coluna(coluna)

    # Retorna a letra normalizada
    return letra


def obter_ultima_coluna(
    planilha: Worksheet,
    ignorar_formatacao: bool = True,
) -> int:
    """
    Retorna o índice da última coluna que contém algum valor.

    Args:
        planilha (Worksheet):
            Planilha que será analisada.
        ignorar_formatacao (bool):
            Quando True, considera somente células com valor.
            Quando False, utiliza o max_column calculado pelo openpyxl.

    Returns:
        int:
            Índice da última coluna utilizada. Retorna 0 quando a planilha
            não possui valores e ignorar_formatacao é True.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Valida a opção de análise da formatação
    _validar_booleano(ignorar_formatacao, "ignorar_formatacao")

    # Retorna diretamente o limite calculado pelo openpyxl quando solicitado
    if not ignorar_formatacao:
        return planilha.max_column

    # Percorre as colunas da direita para a esquerda
    for indice_coluna in range(planilha.max_column, 0, -1):
        # Percorre as linhas da coluna atual
        for indice_linha in range(1, planilha.max_row + 1):
            # Verifica se a célula possui algum valor
            if planilha.cell(indice_linha, indice_coluna).value is not None:
                # Retorna a primeira coluna preenchida encontrada
                return indice_coluna

    # Retorna zero quando nenhuma célula possui valor
    return 0


def obter_valores_coluna(
    planilha: Worksheet,
    coluna: str | int,
    linha_inicial: int = 1,
    linha_final: int | None = None,
    ignorar_vazios: bool = False,
) -> list[Any]:
    """
    Obtém os valores de uma coluna dentro de um intervalo de linhas.

    Args:
        planilha (Worksheet):
            Planilha que contém os dados.
        coluna (str | int):
            Letra ou índice da coluna consultada.
        linha_inicial (int):
            Primeira linha incluída na leitura.
        linha_final (int | None):
            Última linha incluída. Quando None, utiliza max_row.
        ignorar_vazios (bool):
            Define se valores None serão removidos da lista.

    Returns:
        list[Any]:
            Lista com os valores encontrados na coluna.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice numérico da coluna
    indice_coluna = obter_indice_coluna(coluna)

    # Valida a primeira linha da leitura
    inicio = _validar_inteiro(
        linha_inicial,
        "linha_inicial",
        minimo=1,
    )

    # Define a última linha com base na planilha quando não informada
    fim = (
        planilha.max_row
        if linha_final is None
        else _validar_inteiro(linha_final, "linha_final", minimo=1)
    )

    # Verifica se o intervalo de linhas está em ordem crescente
    if fim < inicio:
        # Interrompe a execução quando o intervalo é inválido
        raise ValueError(
            "A linha final não pode ser menor que a linha inicial."
        )

    # Valida a opção de remoção dos valores vazios
    _validar_booleano(ignorar_vazios, "ignorar_vazios")

    # Cria a lista com os valores presentes no intervalo
    valores = [
        planilha.cell(indice_linha, indice_coluna).value
        for indice_linha in range(inicio, fim + 1)
    ]

    # Remove valores None quando solicitado
    if ignorar_vazios:
        valores = [
            valor
            for valor in valores
            if valor is not None
        ]

    # Retorna a lista de valores da coluna
    return valores


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE ALTERAÇÃO
# ----------------------------------------------------------------------------

def definir_valores_coluna(
    planilha: Worksheet,
    coluna: str | int,
    valores: Iterable[Any],
    linha_inicial: int = 1,
) -> Worksheet:
    """
    Preenche uma coluna com os valores de um iterável.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Letra ou índice da coluna de destino.
        valores (Iterable[Any]):
            Valores que serão gravados verticalmente.
        linha_inicial (int):
            Linha onde a gravação será iniciada.

    Returns:
        Worksheet:
            A própria planilha alterada.

    Raises:
        TypeError:
            Caso valores não seja um iterável válido.
        ValueError:
            Caso o iterável esteja vazio.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice numérico da coluna de destino
    indice_coluna = obter_indice_coluna(coluna)

    # Valida a linha inicial da gravação
    inicio = _validar_inteiro(
        linha_inicial,
        "linha_inicial",
        minimo=1,
    )

    # Rejeita strings para evitar a gravação de um caractere por linha
    if isinstance(valores, (str, bytes)) or not isinstance(valores, Iterable):
        # Interrompe a execução quando os valores não formam um iterável válido
        raise TypeError(
            "O parâmetro 'valores' deve ser um iterável não textual."
        )

    # Converte o iterável para lista para verificar seu conteúdo
    lista_valores = list(valores)

    # Verifica se algum valor foi informado
    if not lista_valores:
        # Interrompe a execução quando a lista está vazia
        raise ValueError(
            "O parâmetro 'valores' não pode estar vazio."
        )

    # Percorre os valores mantendo a ordem original
    for deslocamento, valor in enumerate(lista_valores):
        # Grava cada valor na linha correspondente
        planilha.cell(
            row=inicio + deslocamento,
            column=indice_coluna,
            value=valor,
        )

    # Retorna a planilha alterada
    return planilha


def inserir_colunas(
    planilha: Worksheet,
    coluna: str | int,
    quantidade: int = 1,
) -> Worksheet:
    """
    Insere uma ou mais colunas antes da referência informada.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Coluna antes da qual as novas colunas serão inseridas.
        quantidade (int):
            Quantidade de colunas que será inserida.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice da coluna de inserção
    indice_coluna = obter_indice_coluna(coluna)

    # Valida a quantidade de colunas
    quantidade_validada = _validar_inteiro(
        quantidade,
        "quantidade",
        minimo=1,
    )

    # Insere as novas colunas na planilha
    planilha.insert_cols(
        indice_coluna,
        amount=quantidade_validada,
    )

    # Retorna a planilha alterada
    return planilha


def remover_colunas(
    planilha: Worksheet,
    coluna: str | int,
    quantidade: int = 1,
) -> Worksheet:
    """
    Remove uma ou mais colunas a partir da referência informada.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Primeira coluna que será removida.
        quantidade (int):
            Quantidade de colunas que será removida.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice da primeira coluna que será removida
    indice_coluna = obter_indice_coluna(coluna)

    # Valida a quantidade de colunas
    quantidade_validada = _validar_inteiro(
        quantidade,
        "quantidade",
        minimo=1,
    )

    # Remove as colunas selecionadas
    planilha.delete_cols(
        indice_coluna,
        amount=quantidade_validada,
    )

    # Retorna a planilha alterada
    return planilha


def copiar_coluna(
    planilha: Worksheet,
    coluna_origem: str | int,
    coluna_destino: str | int,
    linha_inicial: int = 1,
    linha_final: int | None = None,
    copiar_estilos: bool = True,
    copiar_largura: bool = True,
) -> Worksheet:
    """
    Copia valores, estilos e largura de uma coluna para outra.

    Args:
        planilha (Worksheet):
            Planilha que contém as colunas.
        coluna_origem (str | int):
            Coluna de onde os dados serão copiados.
        coluna_destino (str | int):
            Coluna que receberá os dados.
        linha_inicial (int):
            Primeira linha incluída na cópia.
        linha_final (int | None):
            Última linha incluída. Quando None, utiliza max_row.
        copiar_estilos (bool):
            Define se os estilos das células serão copiados.
        copiar_largura (bool):
            Define se a largura da coluna será copiada.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Normaliza as referências das colunas
    indice_origem, letra_origem = normalizar_coluna(coluna_origem)
    indice_destino, letra_destino = normalizar_coluna(coluna_destino)

    # Verifica se origem e destino representam colunas diferentes
    if indice_origem == indice_destino:
        # Interrompe a execução quando não existe uma coluna de destino distinta
        raise ValueError(
            "As colunas de origem e destino devem ser diferentes."
        )

    # Valida a linha inicial
    inicio = _validar_inteiro(
        linha_inicial,
        "linha_inicial",
        minimo=1,
    )

    # Define e valida a linha final
    fim = (
        planilha.max_row
        if linha_final is None
        else _validar_inteiro(linha_final, "linha_final", minimo=1)
    )

    # Verifica se o intervalo de linhas é válido
    if fim < inicio:
        # Interrompe a execução quando o intervalo está invertido
        raise ValueError(
            "A linha final não pode ser menor que a linha inicial."
        )

    # Valida as opções de cópia
    _validar_booleano(copiar_estilos, "copiar_estilos")
    _validar_booleano(copiar_largura, "copiar_largura")

    # Percorre todas as linhas selecionadas
    for indice_linha in range(inicio, fim + 1):
        # Obtém as células de origem e destino
        celula_origem = planilha.cell(indice_linha, indice_origem)
        celula_destino = planilha.cell(indice_linha, indice_destino)

        # Copia o valor da célula original
        celula_destino.value = celula_origem.value

        # Copia os estilos quando solicitado
        if copiar_estilos:
            _copiar_estilo_celula(
                celula_origem,
                celula_destino,
            )

    # Copia a largura da coluna quando solicitado
    if copiar_largura:
        planilha.column_dimensions[letra_destino].width = (
            planilha.column_dimensions[letra_origem].width
        )

    # Retorna a planilha alterada
    return planilha


def mover_coluna(
    planilha: Worksheet,
    coluna_origem: str | int,
    coluna_destino: str | int,
    copiar_estilos: bool = True,
    copiar_largura: bool = True,
) -> Worksheet:
    """
    Move uma coluna completa para outra posição da planilha.

    A função preserva valores e, opcionalmente, estilos e largura.
    A coluna de origem é removida após a captura dos dados.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna_origem (str | int):
            Coluna que será movida.
        coluna_destino (str | int):
            Posição final desejada para a coluna.
        copiar_estilos (bool):
            Define se os estilos serão preservados.
        copiar_largura (bool):
            Define se a largura será preservada.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Normaliza as posições de origem e destino
    indice_origem, letra_origem = normalizar_coluna(coluna_origem)
    indice_destino, _ = normalizar_coluna(coluna_destino)

    # Verifica se a coluna já está na posição desejada
    if indice_origem == indice_destino:
        # Retorna a planilha sem realizar alterações
        return planilha

    # Valida as opções de preservação
    _validar_booleano(copiar_estilos, "copiar_estilos")
    _validar_booleano(copiar_largura, "copiar_largura")

    # Armazena os dados e estilos da coluna antes de removê-la
    dados_coluna = []

    # Percorre todas as linhas utilizadas da planilha
    for indice_linha in range(1, planilha.max_row + 1):
        # Obtém a célula da coluna de origem
        celula = planilha.cell(indice_linha, indice_origem)

        # Armazena valor e referências de estilo da célula
        dados_coluna.append(
            {
                "valor": celula.value,
                "font": copy(celula.font),
                "fill": copy(celula.fill),
                "border": copy(celula.border),
                "alignment": copy(celula.alignment),
                "number_format": celula.number_format,
                "protection": copy(celula.protection),
            }
        )

    # Armazena a largura da coluna de origem
    largura_origem = planilha.column_dimensions[letra_origem].width

    # Remove a coluna da posição original
    planilha.delete_cols(indice_origem, amount=1)

    # Ajusta o índice de destino porque a remoção desloca colunas à direita
    if indice_origem < indice_destino:
        indice_destino -= 1

    # Insere uma coluna vazia na posição final
    planilha.insert_cols(indice_destino, amount=1)

    # Obtém a letra definitiva da coluna de destino
    letra_destino = converter_numero_para_coluna(indice_destino)

    # Percorre os dados armazenados da coluna original
    for indice_linha, dados in enumerate(dados_coluna, start=1):
        # Obtém a célula que receberá os dados
        celula_destino = planilha.cell(indice_linha, indice_destino)

        # Restaura o valor original da célula
        celula_destino.value = dados["valor"]

        # Restaura os estilos quando solicitado
        if copiar_estilos:
            celula_destino.font = dados["font"]
            celula_destino.fill = dados["fill"]
            celula_destino.border = dados["border"]
            celula_destino.alignment = dados["alignment"]
            celula_destino.number_format = dados["number_format"]
            celula_destino.protection = dados["protection"]

    # Restaura a largura da coluna quando solicitado
    if copiar_largura:
        planilha.column_dimensions[letra_destino].width = largura_origem

    # Retorna a planilha alterada
    return planilha


def limpar_coluna(
    planilha: Worksheet,
    coluna: str | int,
    linha_inicial: int = 1,
    linha_final: int | None = None,
    limpar_estilos: bool = False,
) -> Worksheet:
    """
    Remove os valores de uma coluna dentro do intervalo informado.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Coluna que terá seus valores removidos.
        linha_inicial (int):
            Primeira linha incluída na limpeza.
        linha_final (int | None):
            Última linha incluída. Quando None, utiliza max_row.
        limpar_estilos (bool):
            Define se os estilos também serão substituídos pelo padrão.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice numérico da coluna
    indice_coluna = obter_indice_coluna(coluna)

    # Valida e define o intervalo de linhas
    inicio = _validar_inteiro(linha_inicial, "linha_inicial", minimo=1)
    fim = (
        planilha.max_row
        if linha_final is None
        else _validar_inteiro(linha_final, "linha_final", minimo=1)
    )

    # Verifica se o intervalo está em ordem crescente
    if fim < inicio:
        # Interrompe a execução quando o intervalo é inválido
        raise ValueError(
            "A linha final não pode ser menor que a linha inicial."
        )

    # Valida a opção de limpeza dos estilos
    _validar_booleano(limpar_estilos, "limpar_estilos")

    # Cria uma célula temporária com o estilo padrão do openpyxl
    from openpyxl.cell.cell import Cell

    # Percorre todas as linhas selecionadas
    for indice_linha in range(inicio, fim + 1):
        # Obtém a célula que será limpa
        celula = planilha.cell(indice_linha, indice_coluna)

        # Remove o conteúdo da célula
        celula.value = None

        # Substitui os estilos pelo padrão quando solicitado
        if limpar_estilos:
            celula_padrao = Cell(planilha, row=indice_linha, column=indice_coluna)
            _copiar_estilo_celula(celula_padrao, celula)

    # Retorna a planilha alterada
    return planilha


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE VISIBILIDADE E LARGURA
# ----------------------------------------------------------------------------

def ocultar_coluna(
    planilha: Worksheet,
    coluna: str | int,
) -> Worksheet:
    """Oculta uma coluna da planilha."""

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém a letra normalizada da coluna
    letra = obter_letra_coluna(coluna)

    # Marca a dimensão da coluna como oculta
    planilha.column_dimensions[letra].hidden = True

    # Retorna a planilha alterada
    return planilha


def exibir_coluna(
    planilha: Worksheet,
    coluna: str | int,
) -> Worksheet:
    """Torna uma coluna oculta novamente visível."""

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém a letra normalizada da coluna
    letra = obter_letra_coluna(coluna)

    # Remove a marcação de coluna oculta
    planilha.column_dimensions[letra].hidden = False

    # Retorna a planilha alterada
    return planilha


def definir_largura_coluna(
    planilha: Worksheet,
    coluna: str | int,
    largura: int | float,
) -> Worksheet:
    """
    Define manualmente a largura de uma coluna.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Coluna que receberá a largura.
        largura (int | float):
            Largura positiva da coluna.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém a letra normalizada da coluna
    letra = obter_letra_coluna(coluna)

    # Verifica se a largura é numérica e impede booleanos
    if isinstance(largura, bool) or not isinstance(largura, (int, float)):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'largura' deve ser numérico."
        )

    # Verifica se a largura é positiva
    if largura <= 0:
        # Interrompe a execução quando a largura é inválida
        raise ValueError(
            "O parâmetro 'largura' deve ser maior que zero."
        )

    # Define a largura da dimensão da coluna
    planilha.column_dimensions[letra].width = largura

    # Retorna a planilha alterada
    return planilha


def obter_largura_coluna(
    planilha: Worksheet,
    coluna: str | int,
) -> float | None:
    """Retorna a largura configurada para uma coluna."""

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém a letra normalizada da coluna
    letra = obter_letra_coluna(coluna)

    # Retorna a largura armazenada na dimensão da coluna
    return planilha.column_dimensions[letra].width


def ajustar_largura_coluna(
    planilha: Worksheet,
    coluna: str | int,
    margem: int | float = 2,
    largura_minima: int | float = 1,
    largura_maxima: int | float = 80,
    linha_inicial: int = 1,
    linha_final: int | None = None,
) -> float:
    """
    Ajusta a largura de uma coluna ao maior conteúdo encontrado.

    Args:
        planilha (Worksheet):
            Planilha que será analisada.
        coluna (str | int):
            Coluna que terá sua largura ajustada.
        margem (int | float):
            Espaço adicional aplicado ao maior conteúdo.
        largura_minima (int | float):
            Menor largura permitida.
        largura_maxima (int | float):
            Maior largura permitida.
        linha_inicial (int):
            Primeira linha incluída na análise.
        linha_final (int | None):
            Última linha incluída. Quando None, utiliza max_row.

    Returns:
        float:
            Largura final aplicada à coluna.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice e a letra da coluna
    indice_coluna, letra_coluna = normalizar_coluna(coluna)

    # Valida os parâmetros numéricos de largura
    for nome, valor in {
        "margem": margem,
        "largura_minima": largura_minima,
        "largura_maxima": largura_maxima,
    }.items():
        # Verifica se cada valor é numérico e impede booleanos
        if isinstance(valor, bool) or not isinstance(valor, (int, float)):
            # Interrompe a execução quando um valor possui tipo inválido
            raise TypeError(
                f"O parâmetro '{nome}' deve ser numérico."
            )

    # Verifica se as larguras e a margem respeitam seus limites
    if margem < 0:
        raise ValueError("O parâmetro 'margem' não pode ser negativo.")
    if largura_minima <= 0:
        raise ValueError("A largura mínima deve ser maior que zero.")
    if largura_maxima < largura_minima:
        raise ValueError(
            "A largura máxima não pode ser menor que a largura mínima."
        )

    # Valida e define o intervalo de linhas
    inicio = _validar_inteiro(linha_inicial, "linha_inicial", minimo=1)
    fim = (
        planilha.max_row
        if linha_final is None
        else _validar_inteiro(linha_final, "linha_final", minimo=1)
    )

    # Verifica se o intervalo está em ordem crescente
    if fim < inicio:
        raise ValueError(
            "A linha final não pode ser menor que a linha inicial."
        )

    # Inicia o maior comprimento com zero
    maior_comprimento = 0

    # Percorre as células da coluna dentro do intervalo
    for indice_linha in range(inicio, fim + 1):
        # Obtém o valor da célula atual
        valor = planilha.cell(indice_linha, indice_coluna).value

        # Ignora células vazias durante o cálculo
        if valor is None:
            continue

        # Converte o valor para texto e calcula o maior trecho entre quebras
        comprimento = max(
            len(parte)
            for parte in str(valor).splitlines() or [""]
        )

        # Atualiza o maior comprimento encontrado
        maior_comprimento = max(
            maior_comprimento,
            comprimento,
        )

    # Calcula a largura aplicando margem e limites
    largura_calculada = max(
        largura_minima,
        min(maior_comprimento + margem, largura_maxima),
    )

    # Aplica a largura calculada à coluna
    planilha.column_dimensions[letra_coluna].width = largura_calculada

    # Retorna a largura aplicada
    return float(largura_calculada)


def ajustar_larguras_colunas(
    planilha: Worksheet,
    colunas: Iterable[str | int] | None = None,
    margem: int | float = 2,
    largura_minima: int | float = 1,
    largura_maxima: int | float = 80,
) -> dict[str, float]:
    """
    Ajusta automaticamente a largura de várias colunas.

    Args:
        planilha (Worksheet):
            Planilha que será analisada.
        colunas (Iterable[str | int] | None):
            Colunas que serão ajustadas. Quando None, utiliza todas as
            colunas preenchidas da planilha.
        margem (int | float):
            Espaço adicional aplicado aos conteúdos.
        largura_minima (int | float):
            Menor largura permitida.
        largura_maxima (int | float):
            Maior largura permitida.

    Returns:
        dict[str, float]:
            Dicionário com as letras e as larguras aplicadas.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Define todas as colunas preenchidas quando nenhuma seleção é informada
    if colunas is None:
        ultima_coluna = obter_ultima_coluna(planilha)
        referencias = list(range(1, ultima_coluna + 1))
    else:
        # Rejeita strings para evitar a leitura de uma letra por item
        if isinstance(colunas, (str, bytes)) or not isinstance(colunas, Iterable):
            raise TypeError(
                "O parâmetro 'colunas' deve ser um iterável não textual ou None."
            )

        # Converte o iterável para lista
        referencias = list(colunas)

    # Cria o dicionário que armazenará os resultados
    larguras: dict[str, float] = {}

    # Percorre as colunas selecionadas
    for referencia in referencias:
        # Obtém a letra normalizada da coluna
        letra = obter_letra_coluna(referencia)

        # Ajusta a largura da coluna atual
        largura = ajustar_largura_coluna(
            planilha,
            referencia,
            margem=margem,
            largura_minima=largura_minima,
            largura_maxima=largura_maxima,
        )

        # Armazena a largura aplicada no resultado
        larguras[letra] = largura

    # Retorna o mapeamento de larguras aplicadas
    return larguras


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE CABEÇALHOS
# ----------------------------------------------------------------------------

def listar_colunas_preenchidas(
    planilha: Worksheet,
    linha_cabecalho: int = 1,
    incluir_vazias: bool = False,
) -> list[dict[str, Any]]:
    """
    Lista as colunas e os valores encontrados na linha de cabeçalho.

    Args:
        planilha (Worksheet):
            Planilha que será consultada.
        linha_cabecalho (int):
            Linha onde os cabeçalhos estão armazenados.
        incluir_vazias (bool):
            Define se colunas com cabeçalho vazio serão incluídas.

    Returns:
        list[dict[str, Any]]:
            Lista com índice, letra e valor de cada coluna.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Valida a linha de cabeçalho
    linha = _validar_inteiro(
        linha_cabecalho,
        "linha_cabecalho",
        minimo=1,
    )

    # Valida a opção de inclusão de colunas vazias
    _validar_booleano(incluir_vazias, "incluir_vazias")

    # Obtém a última coluna utilizada da planilha
    ultima_coluna = obter_ultima_coluna(planilha)

    # Cria a lista de resultados
    resultado: list[dict[str, Any]] = []

    # Percorre as colunas utilizadas da planilha
    for indice_coluna in range(1, ultima_coluna + 1):
        # Obtém o valor localizado na linha de cabeçalho
        valor = planilha.cell(linha, indice_coluna).value

        # Ignora cabeçalhos vazios quando solicitado
        if valor is None and not incluir_vazias:
            continue

        # Adiciona os dados da coluna ao resultado
        resultado.append(
            {
                "indice": indice_coluna,
                "letra": converter_numero_para_coluna(indice_coluna),
                "cabecalho": valor,
            }
        )

    # Retorna a lista de colunas encontradas
    return resultado


def localizar_coluna_por_cabecalho(
    planilha: Worksheet,
    cabecalho: str,
    linha_cabecalho: int = 1,
    considerar_maiusculas: bool = False,
    correspondencia_exata: bool = True,
) -> int:
    """
    Localiza o índice de uma coluna por seu cabeçalho.

    Args:
        planilha (Worksheet):
            Planilha que será consultada.
        cabecalho (str):
            Texto procurado na linha de cabeçalho.
        linha_cabecalho (int):
            Linha onde os cabeçalhos estão armazenados.
        considerar_maiusculas (bool):
            Define se a comparação diferencia maiúsculas e minúsculas.
        correspondencia_exata (bool):
            Define se o texto deve ser exatamente igual ao cabeçalho.

    Returns:
        int:
            Índice numérico da coluna encontrada.

    Raises:
        ColunaNaoEncontradaError:
            Caso nenhum cabeçalho corresponda à pesquisa.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Valida o texto procurado
    texto_procurado = _validar_texto(
        cabecalho,
        "cabecalho",
    )

    # Valida a linha onde os cabeçalhos estão armazenados
    linha = _validar_inteiro(
        linha_cabecalho,
        "linha_cabecalho",
        minimo=1,
    )

    # Valida as opções de comparação
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")
    _validar_booleano(correspondencia_exata, "correspondencia_exata")

    # Percorre os cabeçalhos preenchidos da planilha
    for dados_coluna in listar_colunas_preenchidas(
        planilha,
        linha_cabecalho=linha,
        incluir_vazias=False,
    ):
        # Converte o valor do cabeçalho atual para texto
        texto_atual = str(dados_coluna["cabecalho"])

        # Normaliza as letras quando a comparação não diferencia maiúsculas
        if not considerar_maiusculas:
            texto_atual = texto_atual.casefold()
            esperado = texto_procurado.casefold()
        else:
            esperado = texto_procurado

        # Compara o cabeçalho de forma exata ou parcial
        corresponde = (
            texto_atual == esperado
            if correspondencia_exata
            else esperado in texto_atual
        )

        # Retorna o índice quando o cabeçalho corresponde
        if corresponde:
            return dados_coluna["indice"]

    # Interrompe a execução quando nenhum cabeçalho é encontrado
    raise ColunaNaoEncontradaError(
        f"O cabeçalho '{texto_procurado}' não foi encontrado."
    )


def renomear_cabecalho_coluna(
    planilha: Worksheet,
    coluna: str | int,
    novo_cabecalho: Any,
    linha_cabecalho: int = 1,
) -> Worksheet:
    """
    Altera o valor do cabeçalho de uma coluna.

    Args:
        planilha (Worksheet):
            Planilha que será alterada.
        coluna (str | int):
            Coluna cujo cabeçalho será alterado.
        novo_cabecalho (Any):
            Novo valor atribuído à célula de cabeçalho.
        linha_cabecalho (int):
            Linha onde o cabeçalho está armazenado.

    Returns:
        Worksheet:
            A própria planilha alterada.
    """

    # Valida a planilha recebida
    _validar_planilha(planilha)

    # Obtém o índice da coluna do cabeçalho
    indice_coluna = obter_indice_coluna(coluna)

    # Valida a linha do cabeçalho
    linha = _validar_inteiro(
        linha_cabecalho,
        "linha_cabecalho",
        minimo=1,
    )

    # Substitui o valor da célula de cabeçalho
    planilha.cell(
        row=linha,
        column=indice_coluna,
        value=novo_cabecalho,
    )

    # Retorna a planilha alterada
    return planilha


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

__all__ = [
    "ExcelColunasError",
    "PlanilhaInvalidaError",
    "ColunaInvalidaError",
    "ColunaNaoEncontradaError",
    "OperacaoColunaError",
    "converter_coluna_para_numero",
    "converter_numero_para_coluna",
    "normalizar_coluna",
    "obter_indice_coluna",
    "obter_letra_coluna",
    "obter_ultima_coluna",
    "obter_valores_coluna",
    "definir_valores_coluna",
    "inserir_colunas",
    "remover_colunas",
    "copiar_coluna",
    "mover_coluna",
    "limpar_coluna",
    "ocultar_coluna",
    "exibir_coluna",
    "definir_largura_coluna",
    "obter_largura_coluna",
    "ajustar_largura_coluna",
    "ajustar_larguras_colunas",
    "listar_colunas_preenchidas",
    "localizar_coluna_por_cabecalho",
    "renomear_cabecalho_coluna",
]
