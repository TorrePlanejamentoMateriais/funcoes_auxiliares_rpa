"""
============================================================
Modulo: EXCEL / modelo.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 29/09/2026
Ultima Alteracao: 29/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para criar, estruturar, copiar e
    validar modelos padronizados de arquivos Excel utilizando openpyxl.
    O modulo permite definir planilhas, cabecalhos, linhas de exemplo,
    tabelas estruturadas, listas de validacao, metadados e configuracoes
    basicas de visualizacao.
Funcoes disponiveis:
    - criar_workbook_modelo()
    - criar_planilha_modelo()
    - adicionar_cabecalhos()
    - adicionar_linha_exemplo()
    - adicionar_linhas_exemplo()
    - adicionar_tabela_modelo()
    - adicionar_lista_validacao()
    - definir_metadados_modelo()
    - aplicar_configuracoes_modelo()
    - criar_modelo_por_esquema()
    - copiar_planilha_modelo()
    - obter_estrutura_modelo()
    - validar_estrutura_modelo()
    - salvar_modelo()
    - carregar_modelo()
Dependencias:
    - openpyxl
    - pathlib
Historico:
    v1.0.0 - 29/09/2026
        - Criacao inicial do modulo.
        - Inclusao da criacao de Workbooks e planilhas modelo.
        - Inclusao da criacao de cabecalhos, exemplos e tabelas.
        - Inclusao de listas de validacao e metadados.
        - Inclusao da validacao estrutural e persistencia dos modelos.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa copy para duplicar estilos sem compartilhar objetos mutaveis
from copy import copy

# Importa Path para normalizacao e manipulacao dos caminhos
from pathlib import Path

# Importa Iterable e Mapping para validacao das colecoes recebidas
from collections.abc import Iterable, Mapping

# Importa Any para anotacoes de valores flexiveis
from typing import Any

# Importa os recursos centrais do openpyxl
from openpyxl import Workbook, load_workbook

# Importa Cell para anotacao das celulas retornadas
from openpyxl.cell.cell import Cell

# Importa estilos utilizados na formatacao padrao dos modelos
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Importa classes utilizadas nas tabelas estruturadas
from openpyxl.worksheet.table import Table, TableStyleInfo

# Importa DataValidation para criar listas suspensas
from openpyxl.worksheet.datavalidation import DataValidation

# Importa utilitarios para referencias e letras de colunas
from openpyxl.utils import get_column_letter, range_boundaries

# Importa Worksheet para validacao dos objetos de planilha
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define o nome padrao da primeira planilha de um modelo
NOME_PLANILHA_PADRAO = "Dados"

# Define a cor de fundo padrao para os cabecalhos
COR_CABECALHO_PADRAO = "FF1F4E78"

# Define a cor de texto padrao para os cabecalhos
COR_TEXTO_CABECALHO_PADRAO = "FFFFFFFF"

# Define a cor padrao das bordas
COR_BORDA_PADRAO = "FFD9E1F2"

# Define a fonte padrao dos modelos
FONTE_PADRAO = "Calibri"

# Define o tamanho padrao da fonte
TAMANHO_FONTE_PADRAO = 11

# Define a altura padrao da linha de cabecalho
ALTURA_CABECALHO_PADRAO = 24

# Define a largura inicial aplicada as colunas
LARGURA_COLUNA_PADRAO = 15

# Define o estilo padrao das tabelas estruturadas
ESTILO_TABELA_PADRAO = "TableStyleMedium2"

# Define as extensoes permitidas para modelos Excel
EXTENSOES_MODELO = {
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm"
}


# ----------------------------------------------------------------------------
# EXCECOES
# ----------------------------------------------------------------------------

class ExcelModeloError(Exception):
    """Erro base para operacoes relacionadas a modelos Excel."""


class WorkbookModeloInvalidoError(ExcelModeloError, TypeError):
    """Indica que o objeto informado nao e um Workbook valido."""


class PlanilhaModeloInvalidaError(ExcelModeloError, TypeError):
    """Indica que o objeto informado nao e uma Worksheet valida."""


class NomeModeloInvalidoError(ExcelModeloError, ValueError):
    """Indica que um nome de planilha, tabela ou modelo e invalido."""


class EstruturaModeloInvalidaError(ExcelModeloError, ValueError):
    """Indica que o esquema ou a estrutura do modelo e invalida."""


class ModeloNaoEncontradoError(ExcelModeloError, FileNotFoundError):
    """Indica que o arquivo de modelo solicitado nao foi encontrado."""


class ModeloExistenteError(ExcelModeloError, FileExistsError):
    """Indica que o arquivo de destino ja existe."""


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
            O proprio valor validado.

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

    # Retorna o texto normalizado
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
        WorkbookModeloInvalidoError:
            Caso o objeto nao seja um Workbook.
    """
    # Verifica o tipo do objeto recebido
    if not isinstance(workbook, Workbook):
        raise WorkbookModeloInvalidoError(
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
        PlanilhaModeloInvalidaError:
            Caso o objeto nao seja uma Worksheet.
    """
    # Verifica o tipo do objeto recebido
    if not isinstance(planilha, Worksheet):
        raise PlanilhaModeloInvalidaError(
            "O objeto fornecido nao e uma planilha do openpyxl."
        )

    # Retorna a planilha validada
    return planilha


def _validar_nome_planilha(
    nome: str
) -> str:
    """
    Valida um nome destinado a uma planilha Excel.

    Args:
        nome (str):
            Nome da planilha que sera validado.

    Returns:
        str:
            Nome validado sem espacos externos.

    Raises:
        NomeModeloInvalidoError:
            Caso o nome seja longo ou possua caracteres proibidos.
    """
    # Valida o texto do nome
    try:
        nome_validado = _validar_texto(
            nome,
            "nome"
        )
    except (TypeError, ValueError) as erro:
        raise NomeModeloInvalidoError(
            "O nome da planilha e invalido."
        ) from erro

    # Verifica o limite definido pelo Excel
    if len(nome_validado) > 31:
        raise NomeModeloInvalidoError(
            "O nome da planilha nao pode ultrapassar 31 caracteres."
        )

    # Verifica os caracteres proibidos
    proibidos = set("[]:*?/\\")
    encontrados = sorted(
        caractere
        for caractere in proibidos
        if caractere in nome_validado
    )

    if encontrados:
        raise NomeModeloInvalidoError(
            f"O nome da planilha possui caracteres invalidos: {encontrados}."
        )

    # Retorna o nome validado
    return nome_validado


def _validar_cabecalhos(
    cabecalhos: Iterable[Any]
) -> list[str]:
    """
    Valida uma colecao de cabecalhos do modelo.

    Args:
        cabecalhos (Iterable[Any]):
            Colecao com os cabecalhos da planilha.

    Returns:
        list[str]:
            Lista de cabecalhos textuais validados.

    Raises:
        EstruturaModeloInvalidaError:
            Caso a colecao esteja vazia ou possua duplicidades.
    """
    # Rejeita strings isoladas e valores nao iteraveis
    if isinstance(cabecalhos, (str, bytes)) or not isinstance(
        cabecalhos,
        Iterable
    ):
        raise EstruturaModeloInvalidaError(
            "O parametro 'cabecalhos' deve ser um iteravel nao textual."
        )

    # Converte os valores para textos normalizados
    resultado = [
        _validar_texto(
            str(cabecalho),
            "cabecalho"
        )
        for cabecalho in cabecalhos
    ]

    # Rejeita colecoes vazias
    if not resultado:
        raise EstruturaModeloInvalidaError(
            "O modelo deve possuir ao menos um cabecalho."
        )

    # Rejeita nomes duplicados
    if len(resultado) != len(set(resultado)):
        raise EstruturaModeloInvalidaError(
            "Os cabecalhos do modelo nao podem possuir duplicidades."
        )

    # Retorna os cabecalhos validados
    return resultado


def _normalizar_caminho_modelo(
    caminho: str | Path
) -> Path:
    """
    Normaliza e valida o caminho destinado ao modelo Excel.

    Args:
        caminho (str | Path):
            Caminho do arquivo de modelo.

    Returns:
        Path:
            Caminho absoluto com extensao permitida.

    Raises:
        TypeError:
            Caso o caminho nao seja string ou Path.
        NomeModeloInvalidoError:
            Caso a extensao nao seja permitida.
    """
    # Verifica o tipo do caminho
    if not isinstance(caminho, (str, Path)):
        raise TypeError(
            "O parametro 'caminho' deve ser uma string ou Path."
        )

    # Rejeita strings vazias
    if isinstance(caminho, str) and not caminho.strip():
        raise NomeModeloInvalidoError(
            "O caminho do modelo nao pode estar vazio."
        )

    # Normaliza o caminho
    arquivo = Path(caminho).expanduser().resolve()

    # Verifica a extensao do arquivo
    if arquivo.suffix.lower() not in EXTENSOES_MODELO:
        raise NomeModeloInvalidoError(
            f"A extensao do modelo nao e suportada: '{arquivo.suffix}'."
        )

    # Retorna o caminho validado
    return arquivo


# ----------------------------------------------------------------------------
# FUNCOES DE CRIACAO DO MODELO
# ----------------------------------------------------------------------------

def criar_workbook_modelo(
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    remover_planilha_padrao: bool = False
) -> Workbook:
    """
    Cria um Workbook inicial para utilizacao como modelo.

    Args:
        nome_planilha (str):
            Nome da primeira planilha do modelo.
        remover_planilha_padrao (bool):
            Define se o Workbook sera retornado sem planilhas.

    Returns:
        Workbook:
            Workbook configurado para receber a estrutura do modelo.

    Raises:
        NomeModeloInvalidoError:
            Caso o nome da planilha seja invalido.
    """
    # Valida os parametros recebidos
    nome = _validar_nome_planilha(nome_planilha)
    _validar_booleano(
        remover_planilha_padrao,
        "remover_planilha_padrao"
    )

    # Cria o Workbook
    workbook = Workbook()

    # Remove ou renomeia a planilha inicial
    if remover_planilha_padrao:
        workbook.remove(workbook.active)
    else:
        workbook.active.title = nome

    # Configura o calculo automatico
    workbook.calculation.calcMode = "auto"
    workbook.calculation.calcOnSave = True
    workbook.calculation.fullCalcOnLoad = True

    # Retorna o Workbook criado
    return workbook


def criar_planilha_modelo(
    workbook: Workbook,
    nome_planilha: str,
    cabecalhos: Iterable[Any] | None = None,
    indice: int | None = None,
    substituir_existente: bool = False
) -> Worksheet:
    """
    Cria uma planilha no Workbook e pode incluir seus cabecalhos.

    Args:
        workbook (Workbook):
            Workbook que recebera a planilha.
        nome_planilha (str):
            Nome da planilha criada.
        cabecalhos (Iterable[Any] | None):
            Colecao opcional de cabecalhos.
        indice (int | None):
            Posicao opcional da planilha no Workbook.
        substituir_existente (bool):
            Define se uma planilha existente sera removida e recriada.

    Returns:
        Worksheet:
            Planilha criada no Workbook.

    Raises:
        ModeloExistenteError:
            Caso a planilha ja exista e a substituicao esteja desabilitada.
    """
    # Valida o Workbook, nome e opcao de substituicao
    _validar_workbook(workbook)
    nome = _validar_nome_planilha(nome_planilha)
    _validar_booleano(
        substituir_existente,
        "substituir_existente"
    )

    # Valida o indice quando informado
    if indice is not None:
        if isinstance(indice, bool) or not isinstance(indice, int):
            raise TypeError(
                "O parametro 'indice' deve ser inteiro ou None."
            )
        if indice < 0:
            raise ValueError(
                "O parametro 'indice' deve ser maior ou igual a zero."
            )

    # Trata uma planilha existente
    if nome in workbook.sheetnames:
        if not substituir_existente:
            raise ModeloExistenteError(
                f"A planilha '{nome}' ja existe no Workbook."
            )
        workbook.remove(workbook[nome])

    # Cria a planilha na posicao solicitada
    planilha = workbook.create_sheet(
        title=nome,
        index=indice
    )

    # Adiciona os cabecalhos quando informados
    if cabecalhos is not None:
        adicionar_cabecalhos(
            planilha,
            cabecalhos
        )

    # Retorna a planilha criada
    return planilha


def adicionar_cabecalhos(
    planilha: Worksheet,
    cabecalhos: Iterable[Any],
    linha: int = 1,
    coluna_inicial: int = 1,
    formatar: bool = True
) -> list[Cell]:
    """
    Adiciona cabecalhos a uma planilha modelo.

    Args:
        planilha (Worksheet):
            Planilha que recebera os cabecalhos.
        cabecalhos (Iterable[Any]):
            Colecao com os nomes das colunas.
        linha (int):
            Linha que recebera os cabecalhos.
        coluna_inicial (int):
            Primeira coluna utilizada na escrita.
        formatar (bool):
            Define se o estilo padrao sera aplicado.

    Returns:
        list[Cell]:
            Lista com as celulas de cabecalho criadas.

    Raises:
        EstruturaModeloInvalidaError:
            Caso os cabecalhos estejam vazios ou duplicados.
    """
    # Valida a planilha e os cabecalhos
    _validar_planilha(planilha)
    nomes = _validar_cabecalhos(cabecalhos)

    # Valida a linha e a coluna inicial
    for valor, nome in ((linha, "linha"), (coluna_inicial, "coluna_inicial")):
        if isinstance(valor, bool) or not isinstance(valor, int):
            raise TypeError(
                f"O parametro '{nome}' deve ser inteiro."
            )
        if valor < 1:
            raise ValueError(
                f"O parametro '{nome}' deve ser maior ou igual a 1."
            )

    # Valida a opcao de formatacao
    _validar_booleano(
        formatar,
        "formatar"
    )

    # Cria os objetos de estilo utilizados no cabecalho
    fonte = Font(
        name=FONTE_PADRAO,
        size=TAMANHO_FONTE_PADRAO,
        bold=True,
        color=COR_TEXTO_CABECALHO_PADRAO
    )
    preenchimento = PatternFill(
        fill_type="solid",
        fgColor=COR_CABECALHO_PADRAO
    )
    alinhamento = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )
    borda = Border(
        bottom=Side(
            style="thin",
            color=COR_BORDA_PADRAO
        )
    )

    # Cria a lista das celulas alteradas
    celulas: list[Cell] = []

    # Escreve os cabecalhos na planilha
    for deslocamento, nome in enumerate(nomes):
        coluna = coluna_inicial + deslocamento
        celula = planilha.cell(
            row=linha,
            column=coluna,
            value=nome
        )

        # Aplica a formatacao padrao quando solicitada
        if formatar:
            celula.font = copy(fonte)
            celula.fill = copy(preenchimento)
            celula.alignment = copy(alinhamento)
            celula.border = copy(borda)
            planilha.column_dimensions[get_column_letter(coluna)].width = (
                LARGURA_COLUNA_PADRAO
            )

        celulas.append(celula)

    # Ajusta a altura da linha do cabecalho
    if formatar:
        planilha.row_dimensions[linha].height = ALTURA_CABECALHO_PADRAO

    # Retorna as celulas criadas
    return celulas


def adicionar_linha_exemplo(
    planilha: Worksheet,
    valores: Iterable[Any],
    linha: int | None = None,
    coluna_inicial: int = 1
) -> list[Cell]:
    """
    Adiciona uma linha de dados de exemplo a uma planilha modelo.

    Args:
        planilha (Worksheet):
            Planilha que recebera os valores.
        valores (Iterable[Any]):
            Colecao ordenada com os valores de exemplo.
        linha (int | None):
            Linha de destino ou None para adicionar apos a ultima linha.
        coluna_inicial (int):
            Primeira coluna utilizada na escrita.

    Returns:
        list[Cell]:
            Lista com as celulas preenchidas.

    Raises:
        EstruturaModeloInvalidaError:
            Caso a colecao de valores esteja vazia.
    """
    # Valida a planilha e a colecao de valores
    _validar_planilha(planilha)

    if isinstance(valores, (str, bytes)) or not isinstance(valores, Iterable):
        raise EstruturaModeloInvalidaError(
            "O parametro 'valores' deve ser um iteravel nao textual."
        )

    dados = list(valores)

    if not dados:
        raise EstruturaModeloInvalidaError(
            "A linha de exemplo deve possuir ao menos um valor."
        )

    # Define e valida a linha de destino
    linha_destino = planilha.max_row + 1 if linha is None else linha

    for valor, nome in (
        (linha_destino, "linha"),
        (coluna_inicial, "coluna_inicial")
    ):
        if isinstance(valor, bool) or not isinstance(valor, int):
            raise TypeError(
                f"O parametro '{nome}' deve ser inteiro."
            )
        if valor < 1:
            raise ValueError(
                f"O parametro '{nome}' deve ser maior ou igual a 1."
            )

    # Escreve os valores e armazena as celulas criadas
    celulas: list[Cell] = []

    for deslocamento, valor in enumerate(dados):
        celula = planilha.cell(
            row=linha_destino,
            column=coluna_inicial + deslocamento,
            value=valor
        )
        celula.font = Font(
            name=FONTE_PADRAO,
            size=TAMANHO_FONTE_PADRAO
        )
        celula.alignment = Alignment(
            vertical="center"
        )
        celulas.append(celula)

    # Retorna as celulas preenchidas
    return celulas


def adicionar_linhas_exemplo(
    planilha: Worksheet,
    linhas: Iterable[Iterable[Any]],
    linha_inicial: int | None = None
) -> int:
    """
    Adiciona varias linhas de exemplo a uma planilha modelo.

    Args:
        planilha (Worksheet):
            Planilha que recebera os registros.
        linhas (Iterable[Iterable[Any]]):
            Colecao contendo as linhas de exemplo.
        linha_inicial (int | None):
            Primeira linha de destino ou None para continuar a planilha.

    Returns:
        int:
            Quantidade de linhas adicionadas.

    Raises:
        EstruturaModeloInvalidaError:
            Caso linhas nao seja um iteravel valido.
    """
    # Valida a planilha e a colecao principal
    _validar_planilha(planilha)

    if isinstance(linhas, (str, bytes)) or not isinstance(linhas, Iterable):
        raise EstruturaModeloInvalidaError(
            "O parametro 'linhas' deve ser um iteravel de iteraveis."
        )

    registros = list(linhas)

    if not registros:
        return 0

    # Define a primeira linha de destino
    linha_destino = planilha.max_row + 1 if linha_inicial is None else linha_inicial

    if isinstance(linha_destino, bool) or not isinstance(linha_destino, int):
        raise TypeError(
            "O parametro 'linha_inicial' deve ser inteiro ou None."
        )

    if linha_destino < 1:
        raise ValueError(
            "O parametro 'linha_inicial' deve ser maior ou igual a 1."
        )

    # Adiciona cada registro em uma linha consecutiva
    for deslocamento, valores in enumerate(registros):
        adicionar_linha_exemplo(
            planilha,
            valores,
            linha=linha_destino + deslocamento
        )

    # Retorna a quantidade adicionada
    return len(registros)


def adicionar_tabela_modelo(
    planilha: Worksheet,
    nome_tabela: str,
    referencia: str | None = None,
    estilo: str = ESTILO_TABELA_PADRAO,
    exibir_linhas_alternadas: bool = True
) -> Table:
    """
    Adiciona uma tabela estruturada a uma planilha modelo.

    Args:
        planilha (Worksheet):
            Planilha que recebera a tabela.
        nome_tabela (str):
            Nome interno da tabela estruturada.
        referencia (str | None):
            Intervalo da tabela ou None para utilizar a regiao preenchida.
        estilo (str):
            Nome do estilo visual da tabela.
        exibir_linhas_alternadas (bool):
            Define se as linhas alternadas serao exibidas.

    Returns:
        Table:
            Tabela estruturada adicionada a planilha.

    Raises:
        NomeModeloInvalidoError:
            Caso o nome da tabela seja invalido ou duplicado.
    """
    # Valida a planilha, o nome e as opcoes
    _validar_planilha(planilha)
    nome = _validar_texto(
        nome_tabela,
        "nome_tabela"
    )
    estilo_validado = _validar_texto(
        estilo,
        "estilo"
    )
    _validar_booleano(
        exibir_linhas_alternadas,
        "exibir_linhas_alternadas"
    )

    # Valida o nome interno da tabela
    if nome[0].isdigit() or not all(
        caractere.isalnum() or caractere == "_"
        for caractere in nome
    ):
        raise NomeModeloInvalidoError(
            "O nome da tabela deve comecar com letra ou sublinhado e "
            "possuir somente letras, numeros e sublinhado."
        )

    # Rejeita nomes duplicados na planilha
    if nome in planilha.tables:
        raise ModeloExistenteError(
            f"A tabela '{nome}' ja existe na planilha."
        )

    # Define a referencia utilizada pela tabela
    if referencia is None:
        referencia_tabela = (
            f"A1:{get_column_letter(planilha.max_column)}{planilha.max_row}"
        )
    else:
        referencia_tabela = _validar_texto(
            referencia,
            "referencia"
        ).upper()

        try:
            limites = range_boundaries(referencia_tabela)
        except (TypeError, ValueError) as erro:
            raise EstruturaModeloInvalidaError(
                f"A referencia da tabela e invalida: '{referencia}'."
            ) from erro

        if any(limite is None for limite in limites) or min(limites) < 1:
            raise EstruturaModeloInvalidaError(
                f"A referencia da tabela e invalida: '{referencia}'."
            )

    # Cria a tabela e configura seu estilo
    tabela = Table(
        displayName=nome,
        ref=referencia_tabela
    )
    tabela.tableStyleInfo = TableStyleInfo(
        name=estilo_validado,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=exibir_linhas_alternadas,
        showColumnStripes=False
    )

    # Adiciona a tabela a planilha
    planilha.add_table(tabela)

    # Retorna a tabela criada
    return tabela


def adicionar_lista_validacao(
    planilha: Worksheet,
    referencia: str,
    opcoes: Iterable[Any],
    permitir_vazio: bool = True,
    mensagem_erro: str = "Selecione um valor valido."
) -> DataValidation:
    """
    Adiciona uma lista suspensa de validacao a um intervalo.

    Args:
        planilha (Worksheet):
            Planilha que recebera a validacao.
        referencia (str):
            Celula ou intervalo que recebera a lista.
        opcoes (Iterable[Any]):
            Valores exibidos na lista suspensa.
        permitir_vazio (bool):
            Define se celulas vazias serao aceitas.
        mensagem_erro (str):
            Mensagem exibida para valores invalidos.

    Returns:
        DataValidation:
            Objeto de validacao adicionado a planilha.

    Raises:
        EstruturaModeloInvalidaError:
            Caso a lista esteja vazia ou ultrapasse o limite interno.
    """
    # Valida a planilha, a referencia e as opcoes
    _validar_planilha(planilha)
    intervalo = _validar_texto(
        referencia,
        "referencia"
    ).upper()

    try:
        limites = range_boundaries(intervalo)
    except (TypeError, ValueError) as erro:
        raise EstruturaModeloInvalidaError(
            f"A referencia da validacao e invalida: '{referencia}'."
        ) from erro

    if any(limite is None for limite in limites) or min(limites) < 1:
        raise EstruturaModeloInvalidaError(
            f"A referencia da validacao e invalida: '{referencia}'."
        )

    if isinstance(opcoes, (str, bytes)) or not isinstance(opcoes, Iterable):
        raise EstruturaModeloInvalidaError(
            "O parametro 'opcoes' deve ser um iteravel nao textual."
        )

    valores = [
        str(valor)
        for valor in opcoes
    ]

    if not valores:
        raise EstruturaModeloInvalidaError(
            "A lista de validacao deve possuir ao menos uma opcao."
        )

    # Cria a formula textual usada pela lista suspensa
    formula = '"' + ",".join(
        valor.replace('"', '""')
        for valor in valores
    ) + '"'

    # O Excel limita listas diretas a 255 caracteres
    if len(formula) > 255:
        raise EstruturaModeloInvalidaError(
            "A lista direta de validacao nao pode ultrapassar 255 caracteres."
        )

    # Valida as configuracoes complementares
    _validar_booleano(
        permitir_vazio,
        "permitir_vazio"
    )
    mensagem = _validar_texto(
        mensagem_erro,
        "mensagem_erro"
    )

    # Cria e adiciona a validacao
    validacao = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=permitir_vazio,
        showErrorMessage=True,
        error=mensagem,
        errorTitle="Valor invalido"
    )
    validacao.add(intervalo)
    planilha.add_data_validation(validacao)

    # Retorna a validacao criada
    return validacao


def definir_metadados_modelo(
    workbook: Workbook,
    titulo: str,
    assunto: str = "Modelo Excel",
    autor: str = "Automacao",
    descricao: str = ""
) -> Workbook:
    """
    Define os metadados principais de um Workbook modelo.

    Args:
        workbook (Workbook):
            Workbook que recebera os metadados.
        titulo (str):
            Titulo documental do modelo.
        assunto (str):
            Assunto documental do arquivo.
        autor (str):
            Autor registrado nas propriedades.
        descricao (str):
            Descricao opcional do modelo.

    Returns:
        Workbook:
            O proprio Workbook alterado.

    Raises:
        ValueError:
            Caso titulo, assunto ou autor estejam vazios.
    """
    # Valida o Workbook e os textos obrigatorios
    _validar_workbook(workbook)
    titulo_validado = _validar_texto(
        titulo,
        "titulo"
    )
    assunto_validado = _validar_texto(
        assunto,
        "assunto"
    )
    autor_validado = _validar_texto(
        autor,
        "autor"
    )

    # Valida o tipo da descricao opcional
    if not isinstance(descricao, str):
        raise TypeError(
            "O parametro 'descricao' deve ser uma string."
        )

    # Define as propriedades documentais
    workbook.properties.title = titulo_validado
    workbook.properties.subject = assunto_validado
    workbook.properties.creator = autor_validado
    workbook.properties.description = descricao.strip()

    # Retorna o Workbook alterado
    return workbook


def aplicar_configuracoes_modelo(
    planilha: Worksheet,
    congelar_em: str | None = "A2",
    ocultar_linhas_grade: bool = True,
    aplicar_filtro: bool = True,
    linha_cabecalho: int = 1
) -> Worksheet:
    """
    Aplica configuracoes basicas de visualizacao ao modelo.

    Args:
        planilha (Worksheet):
            Planilha que recebera as configuracoes.
        congelar_em (str | None):
            Celula usada no congelamento ou None para desabilitar.
        ocultar_linhas_grade (bool):
            Define se as linhas de grade serao ocultadas.
        aplicar_filtro (bool):
            Define se o filtro automatico sera configurado.
        linha_cabecalho (int):
            Linha que contem os cabecalhos.

    Returns:
        Worksheet:
            A propria planilha alterada.

    Raises:
        EstruturaModeloInvalidaError:
            Caso a referencia de congelamento seja invalida.
    """
    # Valida a planilha e as opcoes
    _validar_planilha(planilha)
    _validar_booleano(
        ocultar_linhas_grade,
        "ocultar_linhas_grade"
    )
    _validar_booleano(
        aplicar_filtro,
        "aplicar_filtro"
    )

    if isinstance(linha_cabecalho, bool) or not isinstance(linha_cabecalho, int):
        raise TypeError(
            "O parametro 'linha_cabecalho' deve ser inteiro."
        )

    if linha_cabecalho < 1:
        raise ValueError(
            "O parametro 'linha_cabecalho' deve ser maior ou igual a 1."
        )

    # Configura o congelamento dos paineis
    if congelar_em is None:
        planilha.freeze_panes = None
    else:
        referencia = _validar_texto(
            congelar_em,
            "congelar_em"
        ).upper()

        try:
            limites = range_boundaries(referencia)
        except (TypeError, ValueError) as erro:
            raise EstruturaModeloInvalidaError(
                f"A referencia de congelamento e invalida: '{congelar_em}'."
            ) from erro

        if any(limite is None for limite in limites) or min(limites) < 1:
            raise EstruturaModeloInvalidaError(
                f"A referencia de congelamento e invalida: '{congelar_em}'."
            )

        if limites[0] != limites[2] or limites[1] != limites[3]:
            raise EstruturaModeloInvalidaError(
                "A referencia de congelamento deve representar uma celula."
            )

        planilha.freeze_panes = referencia

    # Configura as linhas de grade
    planilha.sheet_view.showGridLines = not ocultar_linhas_grade

    # Configura o filtro automatico quando existem colunas
    if aplicar_filtro and planilha.max_column > 0:
        primeira = f"A{linha_cabecalho}"
        ultima = f"{get_column_letter(planilha.max_column)}{linha_cabecalho}"
        planilha.auto_filter.ref = f"{primeira}:{ultima}"
    elif not aplicar_filtro:
        planilha.auto_filter.ref = None

    # Retorna a planilha alterada
    return planilha


def criar_modelo_por_esquema(
    esquema: Mapping[str, Iterable[Any]],
    linhas_exemplo: Mapping[str, Iterable[Iterable[Any]]] | None = None,
    criar_tabelas: bool = False
) -> Workbook:
    """
    Cria um Workbook completo a partir de um esquema de planilhas.

    Args:
        esquema (Mapping[str, Iterable[Any]]):
            Mapeamento entre nomes de planilhas e seus cabecalhos.
        linhas_exemplo (Mapping[str, Iterable[Iterable[Any]]] | None):
            Mapeamento opcional com linhas de exemplo por planilha.
        criar_tabelas (bool):
            Define se tabelas estruturadas serao criadas.

    Returns:
        Workbook:
            Workbook criado conforme o esquema informado.

    Raises:
        EstruturaModeloInvalidaError:
            Caso o esquema esteja vazio ou possua tipo invalido.
    """
    # Valida o esquema principal
    if not isinstance(esquema, Mapping):
        raise EstruturaModeloInvalidaError(
            "O parametro 'esquema' deve ser um mapeamento."
        )

    if not esquema:
        raise EstruturaModeloInvalidaError(
            "O parametro 'esquema' nao pode estar vazio."
        )

    # Valida as linhas de exemplo e a opcao das tabelas
    if linhas_exemplo is not None and not isinstance(linhas_exemplo, Mapping):
        raise EstruturaModeloInvalidaError(
            "O parametro 'linhas_exemplo' deve ser um mapeamento ou None."
        )

    _validar_booleano(
        criar_tabelas,
        "criar_tabelas"
    )

    # Cria um Workbook sem a planilha inicial
    workbook = criar_workbook_modelo(
        remover_planilha_padrao=True
    )

    # Percorre as planilhas definidas no esquema
    for nome_planilha, cabecalhos in esquema.items():
        planilha = criar_planilha_modelo(
            workbook,
            nome_planilha,
            cabecalhos=cabecalhos
        )

        # Adiciona os exemplos da planilha quando informados
        if linhas_exemplo is not None and nome_planilha in linhas_exemplo:
            adicionar_linhas_exemplo(
                planilha,
                linhas_exemplo[nome_planilha]
            )

        # Aplica as configuracoes visuais padrao
        aplicar_configuracoes_modelo(planilha)

        # Cria uma tabela quando existem cabecalhos e dados
        if criar_tabelas and planilha.max_row >= 2:
            nome_tabela = "Tabela" + "".join(
                caractere
                for caractere in str(nome_planilha)
                if caractere.isalnum() or caractere == "_"
            )

            if nome_tabela[0].isdigit():
                nome_tabela = f"Tabela_{nome_tabela}"

            adicionar_tabela_modelo(
                planilha,
                nome_tabela=nome_tabela
            )

    # Retorna o Workbook criado
    return workbook


# ----------------------------------------------------------------------------
# FUNCOES DE COPIA, CONSULTA E VALIDACAO
# ----------------------------------------------------------------------------

def copiar_planilha_modelo(
    planilha_origem: Worksheet,
    workbook_destino: Workbook,
    nome_destino: str | None = None
) -> Worksheet:
    """
    Copia valores, estilos e dimensoes de uma planilha para outro Workbook.

    Args:
        planilha_origem (Worksheet):
            Planilha que sera utilizada como modelo.
        workbook_destino (Workbook):
            Workbook que recebera a copia.
        nome_destino (str | None):
            Nome opcional da nova planilha.

    Returns:
        Worksheet:
            Planilha copiada para o Workbook de destino.

    Raises:
        ModeloExistenteError:
            Caso o nome de destino ja exista.
    """
    # Valida os objetos recebidos
    _validar_planilha(planilha_origem)
    _validar_workbook(workbook_destino)

    # Define e valida o nome da nova planilha
    nome = _validar_nome_planilha(
        planilha_origem.title if nome_destino is None else nome_destino
    )

    if nome in workbook_destino.sheetnames:
        raise ModeloExistenteError(
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

            if celula_origem.hyperlink:
                celula_destino._hyperlink = copy(celula_origem.hyperlink)

            if celula_origem.comment:
                celula_destino.comment = copy(celula_origem.comment)

    # Copia as dimensoes das colunas
    for letra, dimensao in planilha_origem.column_dimensions.items():
        destino.column_dimensions[letra].width = dimensao.width
        destino.column_dimensions[letra].hidden = dimensao.hidden

    # Copia as dimensoes das linhas
    for indice, dimensao in planilha_origem.row_dimensions.items():
        destino.row_dimensions[indice].height = dimensao.height
        destino.row_dimensions[indice].hidden = dimensao.hidden

    # Copia mesclagens e configuracoes visuais
    for intervalo in planilha_origem.merged_cells.ranges:
        destino.merge_cells(str(intervalo))

    destino.freeze_panes = planilha_origem.freeze_panes
    destino.sheet_view.showGridLines = planilha_origem.sheet_view.showGridLines
    destino.auto_filter.ref = planilha_origem.auto_filter.ref

    # Retorna a planilha copiada
    return destino


def obter_estrutura_modelo(
    workbook: Workbook,
    linha_cabecalho: int = 1
) -> dict[str, dict[str, Any]]:
    """
    Retorna a estrutura resumida das planilhas de um Workbook modelo.

    Args:
        workbook (Workbook):
            Workbook que sera analisado.
        linha_cabecalho (int):
            Linha utilizada como cabecalho nas planilhas.

    Returns:
        dict[str, dict[str, Any]]:
            Estrutura com cabecalhos, dimensoes, tabelas e configuracoes.

    Raises:
        ValueError:
            Caso a linha de cabecalho seja invalida.
    """
    # Valida o Workbook e a linha de cabecalho
    _validar_workbook(workbook)

    if isinstance(linha_cabecalho, bool) or not isinstance(linha_cabecalho, int):
        raise TypeError(
            "O parametro 'linha_cabecalho' deve ser inteiro."
        )

    if linha_cabecalho < 1:
        raise ValueError(
            "O parametro 'linha_cabecalho' deve ser maior ou igual a 1."
        )

    # Cria o dicionario de estrutura
    estrutura: dict[str, dict[str, Any]] = {}

    # Analisa cada planilha do Workbook
    for planilha in workbook.worksheets:
        cabecalhos = [
            planilha.cell(
                row=linha_cabecalho,
                column=coluna
            ).value
            for coluna in range(1, planilha.max_column + 1)
        ]

        estrutura[planilha.title] = {
            "cabecalhos": cabecalhos,
            "quantidade_colunas": planilha.max_column,
            "quantidade_linhas": planilha.max_row,
            "tabelas": list(planilha.tables.keys()),
            "congelamento": (
                str(planilha.freeze_panes)
                if planilha.freeze_panes is not None
                else None
            ),
            "filtro": planilha.auto_filter.ref,
            "linhas_grade": planilha.sheet_view.showGridLines
        }

    # Retorna a estrutura encontrada
    return estrutura


def validar_estrutura_modelo(
    workbook: Workbook,
    esquema_esperado: Mapping[str, Iterable[Any]],
    linha_cabecalho: int = 1,
    exigir_ordem: bool = True
) -> dict[str, Any]:
    """
    Valida se um Workbook corresponde ao esquema esperado.

    Args:
        workbook (Workbook):
            Workbook que sera validado.
        esquema_esperado (Mapping[str, Iterable[Any]]):
            Mapeamento entre planilhas e cabecalhos esperados.
        linha_cabecalho (int):
            Linha utilizada como cabecalho.
        exigir_ordem (bool):
            Define se a ordem dos cabecalhos sera considerada.

    Returns:
        dict[str, Any]:
            Resultado detalhado da validacao estrutural.

    Raises:
        EstruturaModeloInvalidaError:
            Caso o esquema esperado seja invalido.
    """
    # Valida o Workbook, esquema e opcoes
    _validar_workbook(workbook)

    if not isinstance(esquema_esperado, Mapping) or not esquema_esperado:
        raise EstruturaModeloInvalidaError(
            "O parametro 'esquema_esperado' deve ser um mapeamento nao vazio."
        )

    if isinstance(linha_cabecalho, bool) or not isinstance(linha_cabecalho, int):
        raise TypeError(
            "O parametro 'linha_cabecalho' deve ser inteiro."
        )

    if linha_cabecalho < 1:
        raise ValueError(
            "O parametro 'linha_cabecalho' deve ser maior ou igual a 1."
        )

    _validar_booleano(
        exigir_ordem,
        "exigir_ordem"
    )

    # Cria as colecoes de divergencias
    planilhas_ausentes: list[str] = []
    cabecalhos_ausentes: dict[str, list[str]] = {}
    cabecalhos_extras: dict[str, list[str]] = {}
    ordem_incorreta: list[str] = []

    # Valida cada planilha esperada
    for nome_planilha, cabecalhos in esquema_esperado.items():
        nome = _validar_nome_planilha(nome_planilha)
        esperados = _validar_cabecalhos(cabecalhos)

        if nome not in workbook.sheetnames:
            planilhas_ausentes.append(nome)
            continue

        planilha = workbook[nome]
        encontrados = [
            str(planilha.cell(linha_cabecalho, coluna).value)
            for coluna in range(1, planilha.max_column + 1)
            if planilha.cell(linha_cabecalho, coluna).value is not None
        ]

        ausentes = [
            cabecalho
            for cabecalho in esperados
            if cabecalho not in encontrados
        ]
        extras = [
            cabecalho
            for cabecalho in encontrados
            if cabecalho not in esperados
        ]

        if ausentes:
            cabecalhos_ausentes[nome] = ausentes
        if extras:
            cabecalhos_extras[nome] = extras
        if exigir_ordem and encontrados != esperados:
            ordem_incorreta.append(nome)

    # Define o resultado final
    valido = not any(
        (
            planilhas_ausentes,
            cabecalhos_ausentes,
            cabecalhos_extras,
            ordem_incorreta
        )
    )

    # Retorna o relatorio da validacao
    return {
        "valido": valido,
        "planilhas_ausentes": planilhas_ausentes,
        "cabecalhos_ausentes": cabecalhos_ausentes,
        "cabecalhos_extras": cabecalhos_extras,
        "ordem_incorreta": ordem_incorreta
    }


# ----------------------------------------------------------------------------
# FUNCOES DE PERSISTENCIA
# ----------------------------------------------------------------------------

def salvar_modelo(
    workbook: Workbook,
    caminho: str | Path,
    sobrescrever: bool = False,
    criar_diretorios: bool = True
) -> Path:
    """
    Salva um Workbook modelo no caminho informado.

    Args:
        workbook (Workbook):
            Workbook que sera salvo.
        caminho (str | Path):
            Caminho de destino do arquivo.
        sobrescrever (bool):
            Define se um arquivo existente podera ser substituido.
        criar_diretorios (bool):
            Define se os diretorios ausentes serao criados.

    Returns:
        Path:
            Caminho absoluto do arquivo salvo.

    Raises:
        ModeloExistenteError:
            Caso o arquivo exista e sobrescrever seja False.
    """
    # Valida o Workbook, caminho e opcoes
    _validar_workbook(workbook)
    arquivo = _normalizar_caminho_modelo(caminho)
    _validar_booleano(
        sobrescrever,
        "sobrescrever"
    )
    _validar_booleano(
        criar_diretorios,
        "criar_diretorios"
    )

    # Rejeita a substituicao nao autorizada
    if arquivo.exists() and not sobrescrever:
        raise ModeloExistenteError(
            f"O arquivo de modelo ja existe: {arquivo}"
        )

    # Cria os diretorios necessarios
    if criar_diretorios:
        arquivo.parent.mkdir(
            parents=True,
            exist_ok=True
        )
    elif not arquivo.parent.exists():
        raise FileNotFoundError(
            f"O diretorio de destino nao foi encontrado: {arquivo.parent}"
        )

    # Salva o Workbook
    workbook.save(arquivo)

    # Retorna o caminho salvo
    return arquivo


def carregar_modelo(
    caminho: str | Path,
    somente_leitura: bool = False,
    preservar_macros: bool | None = None
) -> Workbook:
    """
    Carrega um arquivo Excel para utilizacao como modelo.

    Args:
        caminho (str | Path):
            Caminho do arquivo de modelo.
        somente_leitura (bool):
            Define se o Workbook sera aberto em modo de leitura.
        preservar_macros (bool | None):
            Define a preservacao de macros ou detecta pela extensao.

    Returns:
        Workbook:
            Workbook carregado do arquivo.

    Raises:
        ModeloNaoEncontradoError:
            Caso o arquivo nao exista.
    """
    # Valida o caminho e a opcao de leitura
    arquivo = _normalizar_caminho_modelo(caminho)
    _validar_booleano(
        somente_leitura,
        "somente_leitura"
    )

    # Verifica a existencia do arquivo
    if not arquivo.exists() or not arquivo.is_file():
        raise ModeloNaoEncontradoError(
            f"O arquivo de modelo nao foi encontrado: {arquivo}"
        )

    # Define a preservacao das macros
    if preservar_macros is None:
        manter_vba = arquivo.suffix.lower() in {".xlsm", ".xltm"}
    else:
        manter_vba = _validar_booleano(
            preservar_macros,
            "preservar_macros"
        )

    # Carrega e retorna o Workbook
    return load_workbook(
        arquivo,
        read_only=somente_leitura,
        keep_vba=manter_vba,
        data_only=False
    )


# ----------------------------------------------------------------------------
# EXPORTACOES PUBLICAS
# ----------------------------------------------------------------------------

__all__ = [
    "NOME_PLANILHA_PADRAO",
    "COR_CABECALHO_PADRAO",
    "COR_TEXTO_CABECALHO_PADRAO",
    "COR_BORDA_PADRAO",
    "FONTE_PADRAO",
    "TAMANHO_FONTE_PADRAO",
    "ALTURA_CABECALHO_PADRAO",
    "LARGURA_COLUNA_PADRAO",
    "ESTILO_TABELA_PADRAO",
    "EXTENSOES_MODELO",
    "ExcelModeloError",
    "WorkbookModeloInvalidoError",
    "PlanilhaModeloInvalidaError",
    "NomeModeloInvalidoError",
    "EstruturaModeloInvalidaError",
    "ModeloNaoEncontradoError",
    "ModeloExistenteError",
    "criar_workbook_modelo",
    "criar_planilha_modelo",
    "adicionar_cabecalhos",
    "adicionar_linha_exemplo",
    "adicionar_linhas_exemplo",
    "adicionar_tabela_modelo",
    "adicionar_lista_validacao",
    "definir_metadados_modelo",
    "aplicar_configuracoes_modelo",
    "criar_modelo_por_esquema",
    "copiar_planilha_modelo",
    "obter_estrutura_modelo",
    "validar_estrutura_modelo",
    "salvar_modelo",
    "carregar_modelo"
]
