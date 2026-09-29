"""
============================================================
Módulo: EXCEL / exportacao.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Biblioteca de funções auxiliares para exportar DataFrames, coleções
    de dados e planilhas utilizando pandas e openpyxl. O módulo permite
    criar arquivos, adicionar ou substituir abas, gravar dados em uma
    planilha existente, aplicar formatação básica, criar tabelas
    estruturadas e exportar múltiplos DataFrames.
Funções disponíveis:
    - validar_dataframe()
    - normalizar_nome_planilha()
    - dataframe_para_linhas()
    - exportar_dataframe()
    - exportar_multiplos_dataframes()
    - adicionar_dataframe_planilha()
    - substituir_planilha()
    - exportar_registros()
    - exportar_com_formatacao()
    - exportar_dataframe_como_tabela()
    - obter_dados_exportacao()
Dependências:
    - pandas
    - openpyxl
    - pathlib
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão da exportação simples de DataFrames.
        - Inclusão da exportação de múltiplas planilhas.
        - Inclusão das operações de adicionar e substituir abas.
        - Inclusão de formatação básica e tabela estruturada.
        - Inclusão de validações de caminhos, planilhas e DataFrames.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa Path para manipulação segura dos caminhos de saída
from pathlib import Path

# Importa Any, Iterable e Mapping para anotações de tipos flexíveis
from typing import Any, Iterable, Mapping

# Importa pandas para manipulação e exportação dos DataFrames
import pandas as pd

# Importa recursos do openpyxl para abrir e criar arquivos Excel
from openpyxl import Workbook, load_workbook

# Importa estilos utilizados na formatação básica das planilhas
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Importa classes utilizadas para criar tabelas estruturadas do Excel
from openpyxl.worksheet.table import Table, TableStyleInfo

# Importa utilitário para converter números em letras de colunas
from openpyxl.utils import get_column_letter

# Importa a classe Worksheet para validar objetos de planilha
from openpyxl.worksheet.worksheet import Worksheet


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define as extensões suportadas pelo fluxo de exportação do módulo
EXTENSOES_EXPORTACAO = {
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
}

# Define a extensão padrão dos novos arquivos exportados
EXTENSAO_EXPORTACAO_PADRAO = ".xlsx"

# Define o mecanismo padrão utilizado pelo pandas
ENGINE_EXPORTACAO_PADRAO = "openpyxl"

# Define o nome padrão da planilha de saída
NOME_PLANILHA_PADRAO = "Dados"

# Define o nome padrão utilizado para tabelas estruturadas
NOME_TABELA_PADRAO = "TabelaDados"

# Define o estilo padrão utilizado nas tabelas estruturadas
ESTILO_TABELA_PADRAO = "TableStyleMedium2"

# Define a quantidade máxima de caracteres no nome de uma planilha
LIMITE_NOME_PLANILHA = 31

# Define os caracteres proibidos pelo Excel no nome de uma planilha
CARACTERES_INVALIDOS_PLANILHA = {
    "[",
    "]",
    ":",
    "*",
    "?",
    "/",
    "\\",
}

# Define a cor padrão de preenchimento dos cabeçalhos
COR_CABECALHO_PADRAO = "FF1F4E78"

# Define a cor padrão do texto dos cabeçalhos
COR_TEXTO_CABECALHO_PADRAO = "FFFFFFFF"


# ----------------------------------------------------------------------------
# EXCEÇÕES
# ----------------------------------------------------------------------------

class ExcelExportacaoError(Exception):
    """Erro base para operações de exportação de dados para Excel."""


class DataFrameInvalidoError(ExcelExportacaoError, TypeError):
    """Indica que o objeto informado não é um DataFrame válido."""


class CaminhoExportacaoInvalidoError(ExcelExportacaoError, ValueError):
    """Indica que o caminho ou a extensão de saída é inválida."""


class ArquivoExportacaoExistenteError(ExcelExportacaoError, FileExistsError):
    """Indica que o arquivo de saída já existe e não pode ser sobrescrito."""


class PlanilhaExportacaoError(ExcelExportacaoError):
    """Representa falhas relacionadas à criação ou alteração de planilhas."""


class ExportacaoDataFrameError(ExcelExportacaoError):
    """Indica que o DataFrame não pôde ser exportado para o Excel."""


class TabelaExportacaoError(ExcelExportacaoError):
    """Indica que a tabela estruturada não pôde ser criada."""


# ----------------------------------------------------------------------------
# FUNÇÕES INTERNAS DE VALIDAÇÃO
# ----------------------------------------------------------------------------

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
    minimo: int = 0,
) -> int:
    """
    Valida um inteiro e seu limite mínimo.

    Args:
        valor (int):
            Número inteiro que será validado.
        nome (str):
            Nome do parâmetro utilizado na mensagem de erro.
        minimo (int):
            Menor valor permitido.

    Returns:
        int:
            O número inteiro validado.
    """

    # Impede que valores booleanos sejam interpretados como números inteiros
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

    # Retorna o valor validado
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
            Nome utilizado na mensagem de erro.
        permitir_vazio (bool):
            Define se textos vazios serão aceitos.

    Returns:
        str:
            Texto sem espaços externos.
    """

    # Valida a opção que permite textos vazios
    _validar_booleano(permitir_vazio, "permitir_vazio")

    # Verifica se o valor informado é uma string
    if not isinstance(valor, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            f"O parâmetro '{nome}' deve ser uma string."
        )

    # Remove os espaços externos do texto
    texto = valor.strip()

    # Verifica se o texto obrigatório possui conteúdo
    if not texto and not permitir_vazio:
        # Interrompe a execução quando o texto está vazio
        raise ValueError(
            f"O parâmetro '{nome}' não pode estar vazio."
        )

    # Retorna o texto normalizado
    return texto


def _normalizar_caminho_exportacao(
    caminho: str | Path,
) -> Path:
    """
    Normaliza e valida um caminho de exportação Excel.

    Args:
        caminho (str | Path):
            Caminho do arquivo que será criado ou alterado.

    Returns:
        Path:
            Caminho absoluto com extensão suportada.

    Raises:
        TypeError:
            Caso o caminho não seja string ou Path.
        CaminhoExportacaoInvalidoError:
            Caso a extensão não seja suportada.
    """

    # Verifica se o caminho possui um tipo aceito
    if not isinstance(caminho, (str, Path)):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'caminho' deve ser string ou Path."
        )

    # Verifica se uma string vazia foi informada
    if isinstance(caminho, str) and not caminho.strip():
        # Interrompe a execução quando o caminho está vazio
        raise ValueError(
            "O parâmetro 'caminho' não pode estar vazio."
        )

    # Converte o caminho para sua representação absoluta
    caminho_normalizado = Path(caminho).expanduser().resolve()

    # Verifica se o caminho não possui extensão
    if not caminho_normalizado.suffix:
        # Adiciona a extensão padrão ao arquivo
        caminho_normalizado = caminho_normalizado.with_suffix(
            EXTENSAO_EXPORTACAO_PADRAO
        )

    # Verifica se a extensão é suportada
    if caminho_normalizado.suffix.lower() not in EXTENSOES_EXPORTACAO:
        # Interrompe a execução quando a extensão é inválida
        raise CaminhoExportacaoInvalidoError(
            "A extensão do arquivo não é suportada para exportação: "
            f"'{caminho_normalizado.suffix}'."
        )

    # Retorna o caminho validado
    return caminho_normalizado


def _validar_planilha(planilha: Worksheet) -> Worksheet:
    """Valida se o objeto informado é uma Worksheet do openpyxl."""

    # Verifica se o objeto recebido é uma planilha válida
    if not isinstance(planilha, Worksheet):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'planilha' deve ser uma Worksheet do openpyxl."
        )

    # Retorna a própria planilha validada
    return planilha


def _validar_posicao(
    linha_inicial: int,
    coluna_inicial: int,
) -> tuple[int, int]:
    """Valida a posição inicial utilizada na escrita da planilha."""

    # Valida a linha inicial com índice iniciado em um
    linha = _validar_inteiro(
        linha_inicial,
        "linha_inicial",
        minimo=1,
    )

    # Valida a coluna inicial com índice iniciado em um
    coluna = _validar_inteiro(
        coluna_inicial,
        "coluna_inicial",
        minimo=1,
    )

    # Retorna a posição validada
    return linha, coluna


def _sanitizar_nome_tabela(nome: str) -> str:
    """
    Sanitiza um nome para uso em tabelas estruturadas do Excel.

    Args:
        nome (str):
            Nome solicitado para a tabela.

    Returns:
        str:
            Nome sem espaços e caracteres incompatíveis.
    """

    # Valida o nome recebido
    texto = _validar_texto(nome, "nome_tabela")

    # Substitui caracteres incompatíveis por sublinhado
    resultado = "".join(
        caractere
        if caractere.isalnum() or caractere == "_"
        else "_"
        for caractere in texto
    )

    # Adiciona um prefixo quando o nome começa com número
    if resultado[0].isdigit():
        resultado = f"Tabela_{resultado}"

    # Retorna o nome sanitizado
    return resultado


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE VALIDAÇÃO E CONVERSÃO
# ----------------------------------------------------------------------------

def validar_dataframe(
    dataframe: pd.DataFrame,
    permitir_vazio: bool = True,
    exigir_colunas: bool = True,
) -> pd.DataFrame:
    """
    Valida um DataFrame utilizado em uma exportação Excel.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será validado.
        permitir_vazio (bool):
            Define se DataFrames sem linhas serão aceitos.
        exigir_colunas (bool):
            Define se ao menos uma coluna deve existir.

    Returns:
        pd.DataFrame:
            O próprio DataFrame validado.

    Raises:
        DataFrameInvalidoError:
            Caso o objeto não seja um DataFrame.
        ValueError:
            Caso o DataFrame viole as regras configuradas.
    """

    # Valida as opções booleanas da função
    _validar_booleano(permitir_vazio, "permitir_vazio")
    _validar_booleano(exigir_colunas, "exigir_colunas")

    # Verifica se o objeto recebido é um DataFrame do pandas
    if not isinstance(dataframe, pd.DataFrame):
        # Interrompe a execução quando o objeto é inválido
        raise DataFrameInvalidoError(
            "O objeto fornecido não é um DataFrame do pandas."
        )

    # Verifica se ao menos uma coluna é obrigatória
    if exigir_colunas and len(dataframe.columns) == 0:
        # Interrompe a execução quando não existem colunas
        raise ValueError(
            "O DataFrame deve possuir ao menos uma coluna."
        )

    # Verifica se linhas vazias são permitidas
    if not permitir_vazio and dataframe.empty:
        # Interrompe a execução quando não existem registros
        raise ValueError(
            "O DataFrame não pode estar vazio."
        )

    # Identifica nomes duplicados nas colunas
    colunas_duplicadas = (
        dataframe.columns[
            dataframe.columns.duplicated()
        ]
        .astype(str)
        .unique()
        .tolist()
    )

    # Verifica se existem nomes de colunas duplicados
    if colunas_duplicadas:
        # Interrompe a execução porque o Excel exige cabeçalhos consistentes
        raise ValueError(
            "O DataFrame possui colunas duplicadas: "
            f"{colunas_duplicadas}."
        )

    # Retorna o DataFrame validado
    return dataframe


def normalizar_nome_planilha(
    nome: str,
    substituto: str = "_",
    truncar: bool = True,
) -> str:
    """
    Normaliza um nome para uso em planilhas do Excel.

    Args:
        nome (str):
            Nome solicitado para a planilha.
        substituto (str):
            Texto que substituirá os caracteres inválidos.
        truncar (bool):
            Define se nomes maiores que 31 caracteres serão truncados.

    Returns:
        str:
            Nome compatível com as regras do Excel.
    """

    # Valida o nome original
    nome_validado = _validar_texto(nome, "nome")

    # Verifica se o substituto é textual
    if not isinstance(substituto, str):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'substituto' deve ser uma string."
        )

    # Valida a opção que permite truncamento
    _validar_booleano(truncar, "truncar")

    # Substitui os caracteres proibidos pelo Excel
    nome_normalizado = "".join(
        substituto
        if caractere in CARACTERES_INVALIDOS_PLANILHA
        else caractere
        for caractere in nome_validado
    )

    # Remove apóstrofos das extremidades, que podem gerar nomes problemáticos
    nome_normalizado = nome_normalizado.strip("'").strip()

    # Verifica se o nome ficou vazio após o tratamento
    if not nome_normalizado:
        # Interrompe a execução quando não restou conteúdo válido
        raise ValueError(
            "O nome da planilha ficou vazio após a normalização."
        )

    # Verifica se o nome ultrapassa o limite permitido
    if len(nome_normalizado) > LIMITE_NOME_PLANILHA:
        # Trunca o nome quando essa opção está habilitada
        if truncar:
            nome_normalizado = nome_normalizado[:LIMITE_NOME_PLANILHA]
        else:
            # Interrompe a execução quando o truncamento foi desabilitado
            raise ValueError(
                "O nome da planilha não pode ultrapassar 31 caracteres."
            )

    # Retorna o nome compatível com o Excel
    return nome_normalizado


def dataframe_para_linhas(
    dataframe: pd.DataFrame,
    incluir_cabecalho: bool = True,
    incluir_indice: bool = False,
) -> list[list[Any]]:
    """
    Converte um DataFrame para uma lista de linhas.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será convertido.
        incluir_cabecalho (bool):
            Define se os nomes das colunas serão incluídos.
        incluir_indice (bool):
            Define se o índice do DataFrame será incluído.

    Returns:
        list[list[Any]]:
            Lista bidimensional adequada para escrita em planilhas.
    """

    # Valida o DataFrame recebido
    validar_dataframe(dataframe)

    # Valida as opções de conversão
    _validar_booleano(incluir_cabecalho, "incluir_cabecalho")
    _validar_booleano(incluir_indice, "incluir_indice")

    # Cria uma cópia para evitar alterações no DataFrame original
    dataframe_exportacao = dataframe.copy(deep=True)

    # Inclui o índice como primeira coluna quando solicitado
    if incluir_indice:
        dataframe_exportacao = dataframe_exportacao.reset_index()

    # Cria a lista que armazenará todas as linhas
    linhas: list[list[Any]] = []

    # Adiciona os nomes das colunas quando solicitado
    if incluir_cabecalho:
        linhas.append(
            [str(coluna) for coluna in dataframe_exportacao.columns]
        )

    # Converte valores ausentes do pandas para None e adiciona os registros
    for registro in dataframe_exportacao.itertuples(
        index=False,
        name=None,
    ):
        linhas.append(
            [
                None
                if pd.isna(valor)
                else valor
                for valor in registro
            ]
        )

    # Retorna a matriz criada
    return linhas


# ----------------------------------------------------------------------------
# FUNÇÕES INTERNAS DE ESCRITA E FORMATAÇÃO
# ----------------------------------------------------------------------------

def _escrever_dataframe_planilha(
    planilha: Worksheet,
    dataframe: pd.DataFrame,
    linha_inicial: int = 1,
    coluna_inicial: int = 1,
    incluir_cabecalho: bool = True,
    incluir_indice: bool = False,
) -> tuple[int, int]:
    """
    Escreve um DataFrame em uma Worksheet e retorna a última posição usada.
    """

    # Valida a planilha e o DataFrame
    _validar_planilha(planilha)
    validar_dataframe(dataframe)

    # Valida a posição inicial da escrita
    linha, coluna = _validar_posicao(
        linha_inicial,
        coluna_inicial,
    )

    # Converte o DataFrame para uma matriz de linhas
    linhas = dataframe_para_linhas(
        dataframe,
        incluir_cabecalho=incluir_cabecalho,
        incluir_indice=incluir_indice,
    )

    # Percorre as linhas que serão gravadas
    for deslocamento_linha, valores in enumerate(linhas):
        # Percorre as colunas da linha atual
        for deslocamento_coluna, valor in enumerate(valores):
            # Grava o valor na posição correspondente
            planilha.cell(
                row=linha + deslocamento_linha,
                column=coluna + deslocamento_coluna,
                value=valor,
            )

    # Calcula a última linha utilizada
    ultima_linha = linha + max(len(linhas) - 1, 0)

    # Calcula a quantidade de colunas exportadas
    quantidade_colunas = len(dataframe.columns) + (1 if incluir_indice else 0)

    # Calcula a última coluna utilizada
    ultima_coluna = coluna + max(quantidade_colunas - 1, 0)

    # Retorna os limites finais da exportação
    return ultima_linha, ultima_coluna


def _formatar_cabecalho(
    planilha: Worksheet,
    linha: int,
    coluna_inicial: int,
    coluna_final: int,
    cor_fundo: str = COR_CABECALHO_PADRAO,
    cor_texto: str = COR_TEXTO_CABECALHO_PADRAO,
) -> None:
    """Aplica formatação visual básica ao cabeçalho exportado."""

    # Cria a fonte branca e em negrito
    fonte = Font(
        bold=True,
        color=cor_texto,
    )

    # Cria o preenchimento sólido do cabeçalho
    preenchimento = PatternFill(
        fill_type="solid",
        fgColor=cor_fundo,
    )

    # Cria a borda inferior utilizada para separar o cabeçalho
    borda = Border(
        bottom=Side(
            style="thin",
            color="FF000000",
        )
    )

    # Percorre as células do cabeçalho
    for indice_coluna in range(coluna_inicial, coluna_final + 1):
        # Obtém a célula do cabeçalho atual
        celula = planilha.cell(
            row=linha,
            column=indice_coluna,
        )

        # Aplica a fonte configurada
        celula.font = fonte

        # Aplica o preenchimento configurado
        celula.fill = preenchimento

        # Aplica a borda inferior
        celula.border = borda

        # Centraliza o conteúdo e permite quebra de texto
        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def _ajustar_larguras(
    planilha: Worksheet,
    coluna_inicial: int,
    coluna_final: int,
    linha_inicial: int,
    linha_final: int,
    margem: int = 2,
    largura_maxima: int = 80,
) -> None:
    """Ajusta as larguras das colunas pelo maior conteúdo encontrado."""

    # Percorre todas as colunas exportadas
    for indice_coluna in range(coluna_inicial, coluna_final + 1):
        # Inicia o maior comprimento da coluna
        maior_comprimento = 0

        # Percorre todas as linhas exportadas
        for indice_linha in range(linha_inicial, linha_final + 1):
            # Obtém o valor da célula atual
            valor = planilha.cell(
                row=indice_linha,
                column=indice_coluna,
            ).value

            # Ignora células vazias
            if valor is None:
                continue

            # Calcula o maior trecho entre possíveis quebras de linha
            comprimento = max(
                len(parte)
                for parte in str(valor).splitlines() or [""]
            )

            # Atualiza o maior comprimento da coluna
            maior_comprimento = max(
                maior_comprimento,
                comprimento,
            )

        # Limita a largura calculada ao valor máximo configurado
        largura = min(
            maior_comprimento + margem,
            largura_maxima,
        )

        # Aplica a largura, garantindo ao menos uma unidade
        planilha.column_dimensions[
            get_column_letter(indice_coluna)
        ].width = max(largura, 1)


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS DE EXPORTAÇÃO
# ----------------------------------------------------------------------------

def exportar_dataframe(
    dataframe: pd.DataFrame,
    caminho: str | Path,
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    incluir_indice: bool = False,
    sobrescrever: bool = False,
    criar_diretorios: bool = True,
) -> Path:
    """
    Exporta um DataFrame para um novo arquivo Excel.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.
        caminho (str | Path):
            Caminho do arquivo de saída.
        nome_planilha (str):
            Nome da planilha criada.
        incluir_indice (bool):
            Define se o índice do DataFrame será exportado.
        sobrescrever (bool):
            Define se um arquivo existente pode ser substituído.
        criar_diretorios (bool):
            Define se diretórios inexistentes serão criados.

    Returns:
        Path:
            Caminho absoluto do arquivo criado.
    """

    # Valida o DataFrame e as opções booleanas
    validar_dataframe(dataframe)
    _validar_booleano(incluir_indice, "incluir_indice")
    _validar_booleano(sobrescrever, "sobrescrever")
    _validar_booleano(criar_diretorios, "criar_diretorios")

    # Normaliza o caminho e o nome da planilha
    destino = _normalizar_caminho_exportacao(caminho)
    planilha = normalizar_nome_planilha(nome_planilha)

    # Verifica se o arquivo já existe
    if destino.exists() and not sobrescrever:
        # Interrompe a execução para preservar o arquivo atual
        raise ArquivoExportacaoExistenteError(
            f"O arquivo de destino já existe: {destino}"
        )

    # Cria os diretórios de destino quando solicitado
    if criar_diretorios:
        destino.parent.mkdir(
            parents=True,
            exist_ok=True,
        )
    elif not destino.parent.is_dir():
        # Interrompe a execução quando o diretório obrigatório não existe
        raise FileNotFoundError(
            f"O diretório de destino não existe: {destino.parent}"
        )

    # Executa a exportação através do pandas
    try:
        dataframe.to_excel(
            destino,
            sheet_name=planilha,
            index=incluir_indice,
            engine=ENGINE_EXPORTACAO_PADRAO,
        )
    except Exception as erro:
        # Converte a falha original em uma exceção específica do módulo
        raise ExportacaoDataFrameError(
            f"Não foi possível exportar o DataFrame para: {destino}"
        ) from erro

    # Retorna o caminho do arquivo criado
    return destino


def exportar_multiplos_dataframes(
    dataframes: Mapping[str, pd.DataFrame],
    caminho: str | Path,
    incluir_indice: bool = False,
    sobrescrever: bool = False,
    criar_diretorios: bool = True,
) -> Path:
    """
    Exporta vários DataFrames para planilhas de um único arquivo Excel.

    Args:
        dataframes (Mapping[str, pd.DataFrame]):
            Mapeamento entre nomes de planilhas e DataFrames.
        caminho (str | Path):
            Caminho do arquivo de saída.
        incluir_indice (bool):
            Define se os índices serão exportados.
        sobrescrever (bool):
            Define se um arquivo existente pode ser substituído.
        criar_diretorios (bool):
            Define se diretórios inexistentes serão criados.

    Returns:
        Path:
            Caminho absoluto do arquivo criado.
    """

    # Verifica se foi informado um mapeamento
    if not isinstance(dataframes, Mapping):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'dataframes' deve ser um mapeamento."
        )

    # Verifica se ao menos um DataFrame foi informado
    if not dataframes:
        # Interrompe a execução quando o mapeamento está vazio
        raise ValueError(
            "O parâmetro 'dataframes' não pode estar vazio."
        )

    # Valida as opções da exportação
    _validar_booleano(incluir_indice, "incluir_indice")
    _validar_booleano(sobrescrever, "sobrescrever")
    _validar_booleano(criar_diretorios, "criar_diretorios")

    # Normaliza o caminho de destino
    destino = _normalizar_caminho_exportacao(caminho)

    # Verifica se o arquivo existente pode ser substituído
    if destino.exists() and not sobrescrever:
        # Interrompe a execução para preservar o arquivo atual
        raise ArquivoExportacaoExistenteError(
            f"O arquivo de destino já existe: {destino}"
        )

    # Cria os diretórios necessários
    if criar_diretorios:
        destino.parent.mkdir(
            parents=True,
            exist_ok=True,
        )
    elif not destino.parent.is_dir():
        # Interrompe a execução quando o diretório não existe
        raise FileNotFoundError(
            f"O diretório de destino não existe: {destino.parent}"
        )

    # Cria o conjunto de nomes normalizados para detectar duplicidades
    nomes_utilizados: set[str] = set()

    # Executa a exportação de todas as planilhas
    try:
        with pd.ExcelWriter(
            destino,
            engine=ENGINE_EXPORTACAO_PADRAO,
        ) as writer:
            # Percorre o mapeamento mantendo sua ordem original
            for nome_planilha, dataframe in dataframes.items():
                # Valida cada DataFrame antes da exportação
                validar_dataframe(dataframe)

                # Normaliza o nome da planilha
                nome_normalizado = normalizar_nome_planilha(
                    nome_planilha
                )

                # Verifica duplicidades após a normalização dos nomes
                if nome_normalizado.casefold() in nomes_utilizados:
                    # Interrompe a execução para evitar conflito de planilhas
                    raise PlanilhaExportacaoError(
                        "Dois nomes de planilhas ficaram iguais após a "
                        f"normalização: '{nome_normalizado}'."
                    )

                # Registra o nome normalizado como utilizado
                nomes_utilizados.add(nome_normalizado.casefold())

                # Exporta o DataFrame para sua planilha
                dataframe.to_excel(
                    writer,
                    sheet_name=nome_normalizado,
                    index=incluir_indice,
                )
    except PlanilhaExportacaoError:
        # Mantém a exceção específica gerada durante a validação
        raise
    except Exception as erro:
        # Converte falhas externas em uma exceção do módulo
        raise ExportacaoDataFrameError(
            f"Não foi possível exportar os DataFrames para: {destino}"
        ) from erro

    # Retorna o caminho do arquivo criado
    return destino


def adicionar_dataframe_planilha(
    dataframe: pd.DataFrame,
    caminho: str | Path,
    nome_planilha: str,
    substituir_se_existir: bool = False,
    incluir_indice: bool = False,
) -> Path:
    """
    Adiciona um DataFrame como nova planilha em um arquivo existente.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será adicionado.
        caminho (str | Path):
            Caminho do arquivo Excel existente.
        nome_planilha (str):
            Nome da nova planilha.
        substituir_se_existir (bool):
            Define se uma planilha existente será substituída.
        incluir_indice (bool):
            Define se o índice do DataFrame será exportado.

    Returns:
        Path:
            Caminho absoluto do arquivo alterado.
    """

    # Valida o DataFrame e as opções da operação
    validar_dataframe(dataframe)
    _validar_booleano(
        substituir_se_existir,
        "substituir_se_existir",
    )
    _validar_booleano(incluir_indice, "incluir_indice")

    # Normaliza o caminho e verifica sua existência
    arquivo = _normalizar_caminho_exportacao(caminho)
    if not arquivo.is_file():
        # Interrompe a execução quando o arquivo não existe
        raise FileNotFoundError(
            f"O arquivo Excel não existe: {arquivo}"
        )

    # Normaliza o nome da planilha de destino
    nome_normalizado = normalizar_nome_planilha(nome_planilha)

    # Define o comportamento quando a planilha já existe
    comportamento = (
        "replace"
        if substituir_se_existir
        else "error"
    )

    # Executa a inclusão através do ExcelWriter
    try:
        with pd.ExcelWriter(
            arquivo,
            engine=ENGINE_EXPORTACAO_PADRAO,
            mode="a",
            if_sheet_exists=comportamento,
        ) as writer:
            # Exporta o DataFrame para a nova planilha
            dataframe.to_excel(
                writer,
                sheet_name=nome_normalizado,
                index=incluir_indice,
            )
    except ValueError as erro:
        # Informa de forma clara o conflito de planilhas
        raise PlanilhaExportacaoError(
            f"A planilha '{nome_normalizado}' já existe no arquivo."
        ) from erro
    except Exception as erro:
        # Converte outras falhas em uma exceção do módulo
        raise ExportacaoDataFrameError(
            f"Não foi possível adicionar a planilha '{nome_normalizado}'."
        ) from erro

    # Retorna o caminho do arquivo alterado
    return arquivo


def substituir_planilha(
    dataframe: pd.DataFrame,
    caminho: str | Path,
    nome_planilha: str,
    incluir_indice: bool = False,
    criar_arquivo: bool = False,
) -> Path:
    """
    Substitui uma planilha existente ou cria a planilha quando permitido.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que substituirá a planilha.
        caminho (str | Path):
            Caminho do arquivo Excel.
        nome_planilha (str):
            Nome da planilha substituída.
        incluir_indice (bool):
            Define se o índice será exportado.
        criar_arquivo (bool):
            Define se um novo arquivo será criado quando ele não existir.

    Returns:
        Path:
            Caminho absoluto do arquivo criado ou alterado.
    """

    # Valida o DataFrame e as opções
    validar_dataframe(dataframe)
    _validar_booleano(incluir_indice, "incluir_indice")
    _validar_booleano(criar_arquivo, "criar_arquivo")

    # Normaliza o caminho de destino
    arquivo = _normalizar_caminho_exportacao(caminho)

    # Cria um novo arquivo quando solicitado e necessário
    if not arquivo.exists():
        if criar_arquivo:
            return exportar_dataframe(
                dataframe,
                arquivo,
                nome_planilha=nome_planilha,
                incluir_indice=incluir_indice,
                sobrescrever=False,
                criar_diretorios=True,
            )

        # Interrompe a execução quando o arquivo obrigatório não existe
        raise FileNotFoundError(
            f"O arquivo Excel não existe: {arquivo}"
        )

    # Adiciona ou substitui a planilha existente
    return adicionar_dataframe_planilha(
        dataframe,
        arquivo,
        nome_planilha=nome_planilha,
        substituir_se_existir=True,
        incluir_indice=incluir_indice,
    )


def exportar_registros(
    registros: Iterable[Mapping[str, Any]],
    caminho: str | Path,
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    colunas: Iterable[str] | None = None,
    sobrescrever: bool = False,
) -> Path:
    """
    Exporta uma coleção de dicionários para um arquivo Excel.

    Args:
        registros (Iterable[Mapping[str, Any]]):
            Coleção de registros que será convertida em DataFrame.
        caminho (str | Path):
            Caminho do arquivo de saída.
        nome_planilha (str):
            Nome da planilha criada.
        colunas (Iterable[str] | None):
            Ordem opcional das colunas exportadas.
        sobrescrever (bool):
            Define se um arquivo existente pode ser substituído.

    Returns:
        Path:
            Caminho absoluto do arquivo criado.
    """

    # Rejeita strings para evitar a leitura de caracteres como registros
    if isinstance(registros, (str, bytes)) or not isinstance(
        registros,
        Iterable,
    ):
        # Interrompe a execução quando os registros não formam um iterável
        raise TypeError(
            "O parâmetro 'registros' deve ser um iterável não textual."
        )

    # Converte os registros para lista para permitir múltiplas validações
    lista_registros = list(registros)

    # Verifica se ao menos um registro foi informado
    if not lista_registros:
        # Interrompe a execução quando a coleção está vazia
        raise ValueError(
            "O parâmetro 'registros' não pode estar vazio."
        )

    # Verifica se todos os registros são mapeamentos
    if not all(
        isinstance(registro, Mapping)
        for registro in lista_registros
    ):
        # Interrompe a execução quando algum registro é inválido
        raise TypeError(
            "Todos os registros devem ser mapeamentos."
        )

    # Normaliza a seleção opcional de colunas
    colunas_normalizadas = None
    if colunas is not None:
        # Rejeita strings para evitar uma coluna por caractere
        if isinstance(colunas, (str, bytes)) or not isinstance(
            colunas,
            Iterable,
        ):
            raise TypeError(
                "O parâmetro 'colunas' deve ser um iterável não textual."
            )

        # Converte e valida os nomes das colunas
        colunas_normalizadas = [
            _validar_texto(coluna, "coluna")
            for coluna in colunas
        ]

        # Verifica se a coleção de colunas possui conteúdo
        if not colunas_normalizadas:
            raise ValueError(
                "O parâmetro 'colunas' não pode estar vazio."
            )

    # Cria o DataFrame com os registros recebidos
    dataframe = pd.DataFrame.from_records(
        lista_registros,
        columns=colunas_normalizadas,
    )

    # Exporta o DataFrame criado
    return exportar_dataframe(
        dataframe,
        caminho,
        nome_planilha=nome_planilha,
        incluir_indice=False,
        sobrescrever=sobrescrever,
        criar_diretorios=True,
    )


def exportar_com_formatacao(
    dataframe: pd.DataFrame,
    caminho: str | Path,
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    incluir_indice: bool = False,
    sobrescrever: bool = False,
    congelar_cabecalho: bool = True,
    aplicar_filtro: bool = True,
    ajustar_larguras: bool = True,
    cor_cabecalho: str = COR_CABECALHO_PADRAO,
    cor_texto_cabecalho: str = COR_TEXTO_CABECALHO_PADRAO,
) -> Path:
    """
    Exporta um DataFrame aplicando formatação básica à planilha.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.
        caminho (str | Path):
            Caminho do arquivo de saída.
        nome_planilha (str):
            Nome da planilha criada.
        incluir_indice (bool):
            Define se o índice será exportado.
        sobrescrever (bool):
            Define se um arquivo existente pode ser substituído.
        congelar_cabecalho (bool):
            Define se a primeira linha de dados permanecerá congelada.
        aplicar_filtro (bool):
            Define se o autofiltro será aplicado ao intervalo.
        ajustar_larguras (bool):
            Define se as larguras serão ajustadas automaticamente.
        cor_cabecalho (str):
            Cor ARGB do preenchimento do cabeçalho.
        cor_texto_cabecalho (str):
            Cor ARGB da fonte do cabeçalho.

    Returns:
        Path:
            Caminho absoluto do arquivo formatado.
    """

    # Valida o DataFrame e todas as opções booleanas
    validar_dataframe(dataframe)
    for nome, valor in {
        "incluir_indice": incluir_indice,
        "sobrescrever": sobrescrever,
        "congelar_cabecalho": congelar_cabecalho,
        "aplicar_filtro": aplicar_filtro,
        "ajustar_larguras": ajustar_larguras,
    }.items():
        _validar_booleano(valor, nome)

    # Valida as cores informadas
    cor_fundo = _validar_texto(cor_cabecalho, "cor_cabecalho").upper()
    cor_texto = _validar_texto(
        cor_texto_cabecalho,
        "cor_texto_cabecalho",
    ).upper()

    # Realiza a exportação inicial do DataFrame
    destino = exportar_dataframe(
        dataframe,
        caminho,
        nome_planilha=nome_planilha,
        incluir_indice=incluir_indice,
        sobrescrever=sobrescrever,
        criar_diretorios=True,
    )

    # Normaliza novamente o nome para localizar a planilha criada
    nome_normalizado = normalizar_nome_planilha(nome_planilha)

    # Abre o Workbook para aplicação dos estilos
    workbook = load_workbook(destino)

    try:
        # Obtém a planilha exportada
        planilha = workbook[nome_normalizado]

        # Calcula a quantidade total de colunas gravadas
        quantidade_colunas = len(dataframe.columns) + (
            1
            if incluir_indice
            else 0
        )

        # Calcula a última linha da exportação
        ultima_linha = len(dataframe.index) + 1

        # Aplica a formatação visual ao cabeçalho
        _formatar_cabecalho(
            planilha,
            linha=1,
            coluna_inicial=1,
            coluna_final=quantidade_colunas,
            cor_fundo=cor_fundo,
            cor_texto=cor_texto,
        )

        # Congela a primeira linha quando solicitado
        if congelar_cabecalho:
            planilha.freeze_panes = "A2"

        # Aplica o autofiltro quando existem colunas exportadas
        if aplicar_filtro and quantidade_colunas > 0:
            planilha.auto_filter.ref = (
                f"A1:{get_column_letter(quantidade_colunas)}"
                f"{max(ultima_linha, 1)}"
            )

        # Ajusta as larguras das colunas quando solicitado
        if ajustar_larguras and quantidade_colunas > 0:
            _ajustar_larguras(
                planilha,
                coluna_inicial=1,
                coluna_final=quantidade_colunas,
                linha_inicial=1,
                linha_final=max(ultima_linha, 1),
            )

        # Salva as alterações de formatação
        workbook.save(destino)
    except Exception as erro:
        # Converte a falha em uma exceção específica de formatação/exportação
        raise ExportacaoDataFrameError(
            "O arquivo foi criado, mas a formatação não pôde ser aplicada."
        ) from erro
    finally:
        # Fecha o Workbook sempre que a operação for concluída
        workbook.close()

    # Retorna o caminho do arquivo formatado
    return destino


def exportar_dataframe_como_tabela(
    dataframe: pd.DataFrame,
    caminho: str | Path,
    nome_planilha: str = NOME_PLANILHA_PADRAO,
    nome_tabela: str = NOME_TABELA_PADRAO,
    estilo_tabela: str = ESTILO_TABELA_PADRAO,
    sobrescrever: bool = False,
    incluir_indice: bool = False,
    ajustar_larguras: bool = True,
) -> Path:
    """
    Exporta um DataFrame como tabela estruturada do Excel.

    Args:
        dataframe (pd.DataFrame):
            DataFrame que será exportado.
        caminho (str | Path):
            Caminho do arquivo de saída.
        nome_planilha (str):
            Nome da planilha criada.
        nome_tabela (str):
            Nome interno da tabela estruturada.
        estilo_tabela (str):
            Estilo visual aplicado à tabela.
        sobrescrever (bool):
            Define se um arquivo existente pode ser substituído.
        incluir_indice (bool):
            Define se o índice será exportado.
        ajustar_larguras (bool):
            Define se as larguras serão ajustadas.

    Returns:
        Path:
            Caminho absoluto do arquivo criado.
    """

    # Exige ao menos uma linha para criar uma tabela estruturada útil
    validar_dataframe(
        dataframe,
        permitir_vazio=False,
        exigir_colunas=True,
    )

    # Valida as opções booleanas
    _validar_booleano(sobrescrever, "sobrescrever")
    _validar_booleano(incluir_indice, "incluir_indice")
    _validar_booleano(ajustar_larguras, "ajustar_larguras")

    # Normaliza nomes e estilo
    nome_planilha_normalizado = normalizar_nome_planilha(nome_planilha)
    nome_tabela_normalizado = _sanitizar_nome_tabela(nome_tabela)
    estilo = _validar_texto(estilo_tabela, "estilo_tabela")

    # Exporta o DataFrame sem aplicar autofiltro independente
    destino = exportar_dataframe(
        dataframe,
        caminho,
        nome_planilha=nome_planilha_normalizado,
        incluir_indice=incluir_indice,
        sobrescrever=sobrescrever,
        criar_diretorios=True,
    )

    # Abre o arquivo criado para incluir a tabela estruturada
    workbook = load_workbook(destino)

    try:
        # Obtém a planilha de saída
        planilha = workbook[nome_planilha_normalizado]

        # Calcula o tamanho do intervalo exportado
        quantidade_colunas = len(dataframe.columns) + (
            1
            if incluir_indice
            else 0
        )
        ultima_linha = len(dataframe.index) + 1
        ultima_coluna = get_column_letter(quantidade_colunas)
        referencia = f"A1:{ultima_coluna}{ultima_linha}"

        # Cria a tabela estruturada no intervalo calculado
        tabela = Table(
            displayName=nome_tabela_normalizado,
            ref=referencia,
        )

        # Configura o estilo visual da tabela
        tabela.tableStyleInfo = TableStyleInfo(
            name=estilo,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        # Adiciona a tabela à planilha
        planilha.add_table(tabela)

        # Congela a primeira linha de dados
        planilha.freeze_panes = "A2"

        # Ajusta as larguras quando solicitado
        if ajustar_larguras:
            _ajustar_larguras(
                planilha,
                coluna_inicial=1,
                coluna_final=quantidade_colunas,
                linha_inicial=1,
                linha_final=ultima_linha,
            )

        # Salva as alterações realizadas
        workbook.save(destino)
    except Exception as erro:
        # Converte a falha em uma exceção específica de tabela
        raise TabelaExportacaoError(
            "Não foi possível criar a tabela estruturada no arquivo Excel."
        ) from erro
    finally:
        # Fecha o Workbook sempre que a operação for concluída
        workbook.close()

    # Retorna o caminho do arquivo criado
    return destino


def obter_dados_exportacao(
    caminho: str | Path,
) -> dict[str, Any]:
    """
    Retorna metadados básicos de um arquivo exportado.

    Args:
        caminho (str | Path):
            Caminho do arquivo Excel existente.

    Returns:
        dict[str, Any]:
            Dicionário com caminho, tamanho, planilhas e propriedades.
    """

    # Normaliza o caminho recebido
    arquivo = _normalizar_caminho_exportacao(caminho)

    # Verifica se o arquivo existe
    if not arquivo.is_file():
        # Interrompe a execução quando o arquivo não existe
        raise FileNotFoundError(
            f"O arquivo Excel não existe: {arquivo}"
        )

    # Abre o Workbook em modo de leitura para consultar as planilhas
    workbook = load_workbook(
        arquivo,
        read_only=True,
        data_only=False,
    )

    try:
        # Cria a estrutura de metadados do arquivo
        dados = {
            "nome": arquivo.name,
            "caminho": arquivo,
            "diretorio": arquivo.parent,
            "extensao": arquivo.suffix.lower(),
            "tamanho_bytes": arquivo.stat().st_size,
            "planilhas": list(workbook.sheetnames),
            "quantidade_planilhas": len(workbook.sheetnames),
            "possui_macro": arquivo.suffix.lower() in {".xlsm", ".xltm"},
        }
    finally:
        # Fecha o Workbook após a consulta
        workbook.close()

    # Retorna os metadados coletados
    return dados


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

# Define explicitamente os objetos públicos disponibilizados pelo módulo
__all__ = [
    "EXTENSOES_EXPORTACAO",
    "EXTENSAO_EXPORTACAO_PADRAO",
    "ENGINE_EXPORTACAO_PADRAO",
    "NOME_PLANILHA_PADRAO",
    "NOME_TABELA_PADRAO",
    "ESTILO_TABELA_PADRAO",
    "LIMITE_NOME_PLANILHA",
    "CARACTERES_INVALIDOS_PLANILHA",
    "COR_CABECALHO_PADRAO",
    "COR_TEXTO_CABECALHO_PADRAO",
    "ExcelExportacaoError",
    "DataFrameInvalidoError",
    "CaminhoExportacaoInvalidoError",
    "ArquivoExportacaoExistenteError",
    "PlanilhaExportacaoError",
    "ExportacaoDataFrameError",
    "TabelaExportacaoError",
    "validar_dataframe",
    "normalizar_nome_planilha",
    "dataframe_para_linhas",
    "exportar_dataframe",
    "exportar_multiplos_dataframes",
    "adicionar_dataframe_planilha",
    "substituir_planilha",
    "exportar_registros",
    "exportar_com_formatacao",
    "exportar_dataframe_como_tabela",
    "obter_dados_exportacao",
]
