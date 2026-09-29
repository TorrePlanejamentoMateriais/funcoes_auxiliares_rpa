"""
============================================================
Módulo: EXCEL / constantes.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Centraliza constantes, enumerações, formatos, extensões, cores,
    alinhamentos, tipos de borda e limites utilizados pelos módulos
    auxiliares de automação de arquivos Excel.
Constantes disponíveis:
    - EXTENSOES_EXCEL
    - EXTENSOES_OPENPYXL
    - EXTENSOES_COM_MACRO
    - EXTENSOES_MODELO
    - EXTENSAO_EXCEL_LEGADA
    - LIMITE_LINHAS_EXCEL
    - LIMITE_COLUNAS_EXCEL
    - LIMITE_CARACTERES_CELULA
    - LIMITE_CARACTERES_PLANILHA
    - FORMATOS_DATA
    - FORMATOS_HORA
    - FORMATOS_NUMERO
    - FORMATOS_MOEDA
    - FORMATOS_PERCENTUAL
    - CORES_EXCEL
    - CORES_CORPORATIVAS
    - CARACTERES_INVALIDOS_PLANILHA
    - CARACTERES_INVALIDOS_ARQUIVO
Enumerações disponíveis:
    - TipoArquivoExcel
    - TipoAlinhamentoHorizontal
    - TipoAlinhamentoVertical
    - TipoBorda
    - TipoPreenchimento
    - TipoFormatoCelula
    - TipoOperacaoArquivo
Dependências:
    - enum
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão das extensões suportadas pelo Excel e openpyxl.
        - Inclusão dos limites estruturais do Excel.
        - Inclusão de formatos de datas, números, moedas e percentuais.
        - Inclusão das cores e enumerações utilizadas na biblioteca.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa Enum para criação de conjuntos de valores textuais controlados
from enum import Enum


# ----------------------------------------------------------------------------
# EXTENSÕES DE ARQUIVOS
# ----------------------------------------------------------------------------

# Define as extensões conhecidas de arquivos do Excel
EXTENSOES_EXCEL = {
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xls",
}

# Define as extensões que podem ser manipuladas diretamente pelo openpyxl
EXTENSOES_OPENPYXL = {
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
}

# Define as extensões que podem armazenar projetos VBA
EXTENSOES_COM_MACRO = {
    ".xlsm",
    ".xltm",
}

# Define as extensões utilizadas por modelos do Excel
EXTENSOES_MODELO = {
    ".xltx",
    ".xltm",
}

# Define as extensões utilizadas por pastas de trabalho comuns
EXTENSOES_WORKBOOK = {
    ".xlsx",
    ".xlsm",
}

# Define a extensão legada que não é suportada pelo openpyxl
EXTENSAO_EXCEL_LEGADA = ".xls"

# Define a extensão padrão para novos arquivos Excel
EXTENSAO_EXCEL_PADRAO = ".xlsx"

# Define a extensão padrão para arquivos com macros
EXTENSAO_EXCEL_MACRO = ".xlsm"


# ----------------------------------------------------------------------------
# LIMITES DO EXCEL
# ----------------------------------------------------------------------------

# Define a quantidade máxima de linhas permitida em uma planilha moderna
LIMITE_LINHAS_EXCEL = 1_048_576

# Define a quantidade máxima de colunas permitida em uma planilha moderna
LIMITE_COLUNAS_EXCEL = 16_384

# Define a letra da última coluna disponível no Excel moderno
ULTIMA_COLUNA_EXCEL = "XFD"

# Define o número máximo de caracteres aceito em uma célula
LIMITE_CARACTERES_CELULA = 32_767

# Define o número máximo de caracteres aceito no nome de uma planilha
LIMITE_CARACTERES_PLANILHA = 31

# Define o número máximo de caracteres visíveis em cabeçalhos e rodapés
LIMITE_CARACTERES_CABECALHO_RODAPE = 255

# Define a quantidade máxima de caracteres usada por nomes definidos
LIMITE_CARACTERES_NOME_DEFINIDO = 255

# Define a largura máxima prática de uma coluna no Excel
LARGURA_MAXIMA_COLUNA = 255

# Define a altura máxima prática de uma linha em pontos
ALTURA_MAXIMA_LINHA = 409

# Define a largura padrão utilizada quando nenhuma configuração é informada
LARGURA_PADRAO_COLUNA = 8.43

# Define a altura padrão utilizada pelo Excel em pontos
ALTURA_PADRAO_LINHA = 15.0


# ----------------------------------------------------------------------------
# CARACTERES INVÁLIDOS
# ----------------------------------------------------------------------------

# Define os caracteres que não podem existir no nome de uma planilha
CARACTERES_INVALIDOS_PLANILHA = {
    "[",
    "]",
    ":",
    "*",
    "?",
    "/",
    "\\",
}

# Define os caracteres proibidos em nomes de arquivos no Windows
CARACTERES_INVALIDOS_ARQUIVO = {
    "<",
    ">",
    ":",
    '"',
    "/",
    "\\",
    "|",
    "?",
    "*",
}

# Define os nomes reservados que não devem ser usados como arquivos no Windows
NOMES_RESERVADOS_WINDOWS = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}


# ----------------------------------------------------------------------------
# FORMATOS DE DATA E HORA
# ----------------------------------------------------------------------------

# Agrupa formatos comuns de data utilizados nas células do Excel
FORMATOS_DATA = {
    "data_brasileira": "dd/mm/yyyy",
    "data_brasileira_curta": "dd/mm/yy",
    "data_iso": "yyyy-mm-dd",
    "data_extenso": "dd 'de' mmmm 'de' yyyy",
    "mes_ano": "mm/yyyy",
    "mes_ano_extenso": "mmmm/yyyy",
    "dia_mes": "dd/mm",
    "ano": "yyyy",
}

# Agrupa formatos comuns de hora utilizados nas células do Excel
FORMATOS_HORA = {
    "hora_minuto": "hh:mm",
    "hora_completa": "hh:mm:ss",
    "hora_12": "hh:mm AM/PM",
    "duracao": "[h]:mm:ss",
}

# Agrupa formatos que combinam data e hora
FORMATOS_DATA_HORA = {
    "data_hora": "dd/mm/yyyy hh:mm",
    "data_hora_completa": "dd/mm/yyyy hh:mm:ss",
    "iso_data_hora": "yyyy-mm-dd hh:mm:ss",
}

# Define o formato padrão de data da biblioteca
FORMATO_DATA_PADRAO = FORMATOS_DATA["data_brasileira"]

# Define o formato padrão de data e hora da biblioteca
FORMATO_DATA_HORA_PADRAO = FORMATOS_DATA_HORA["data_hora_completa"]


# ----------------------------------------------------------------------------
# FORMATOS NUMÉRICOS
# ----------------------------------------------------------------------------

# Agrupa formatos utilizados para números inteiros e decimais
FORMATOS_NUMERO = {
    "inteiro": "0",
    "inteiro_milhar": "#,##0",
    "decimal_1": "0.0",
    "decimal_2": "0.00",
    "decimal_4": "0.0000",
    "decimal_milhar_2": "#,##0.00",
    "negativo_vermelho": "#,##0.00;[Red](#,##0.00)",
    "zero_como_traco": '#,##0.00;[Red](#,##0.00);"-"',
}

# Agrupa formatos monetários em moedas comuns
FORMATOS_MOEDA = {
    "real": 'R$ #,##0.00;[Red](R$ #,##0.00)',
    "real_zero_traco": 'R$ #,##0.00;[Red](R$ #,##0.00);"-"',
    "dolar": '$ #,##0.00;[Red]($ #,##0.00)',
    "euro": '€ #,##0.00;[Red](€ #,##0.00)',
    "libra": '£ #,##0.00;[Red](£ #,##0.00)',
}

# Agrupa formatos percentuais com diferentes precisões
FORMATOS_PERCENTUAL = {
    "percentual_inteiro": "0%",
    "percentual_1": "0.0%",
    "percentual_2": "0.00%",
    "percentual_negativo_vermelho": "0.0%;[Red](0.0%)",
}

# Agrupa formatos de texto e identificação
FORMATOS_TEXTO = {
    "texto": "@",
    "cep": "00000-000",
    "cpf": "000.000.000-00",
    "cnpj": "00.000.000/0000-00",
    "telefone": "(00) 0000-0000",
    "celular": "(00) 00000-0000",
}

# Define o formato numérico padrão da biblioteca
FORMATO_NUMERO_PADRAO = FORMATOS_NUMERO["decimal_milhar_2"]

# Define o formato monetário padrão da biblioteca
FORMATO_MOEDA_PADRAO = FORMATOS_MOEDA["real"]

# Define o formato percentual padrão da biblioteca
FORMATO_PERCENTUAL_PADRAO = FORMATOS_PERCENTUAL["percentual_1"]


# ----------------------------------------------------------------------------
# CORES
# ----------------------------------------------------------------------------

# Agrupa cores básicas no padrão hexadecimal ARGB utilizado pelo openpyxl
CORES_EXCEL = {
    "preto": "FF000000",
    "branco": "FFFFFFFF",
    "vermelho": "FFFF0000",
    "verde": "FF008000",
    "azul": "FF0000FF",
    "amarelo": "FFFFFF00",
    "laranja": "FFFFA500",
    "roxo": "FF800080",
    "cinza": "FF808080",
    "cinza_claro": "FFD9E1F2",
    "azul_claro": "FFDDEBF7",
    "verde_claro": "FFE2F0D9",
    "vermelho_claro": "FFF4CCCC",
    "amarelo_claro": "FFFFF2CC",
    "laranja_claro": "FFFCE4D6",
    "roxo_claro": "FFE4DFEC",
    "turquesa": "FF00A6A6",
}

# Agrupa cores corporativas sugeridas para relatórios e automações
CORES_CORPORATIVAS = {
    "azul_escuro": "FF1F4E78",
    "azul_medio": "FF4472C4",
    "azul_claro": "FFD9EAF7",
    "cinza_escuro": "FF404040",
    "cinza_medio": "FF7F7F7F",
    "cinza_claro": "FFF2F2F2",
    "verde_sucesso": "FF70AD47",
    "amarelo_alerta": "FFFFC000",
    "vermelho_erro": "FFC00000",
    "laranja_revisao": "FFED7D31",
    "roxo_controle": "FF7030A0",
    "turquesa_destaque": "FF00A6A6",
}

# Define as cores de texto recomendadas conforme a origem do valor
CORES_TEXTO_MODELO = {
    "entrada_usuario": "FF0000FF",
    "formula": "FF000000",
    "vinculo_interno": "FF008000",
    "vinculo_externo": "FFFF0000",
    "constante": "FF808080",
}

# Define a cor padrão utilizada em cabeçalhos
COR_CABECALHO_PADRAO = CORES_CORPORATIVAS["azul_escuro"]

# Define a cor padrão utilizada no texto dos cabeçalhos
COR_TEXTO_CABECALHO_PADRAO = CORES_EXCEL["branco"]


# ----------------------------------------------------------------------------
# ALINHAMENTOS, BORDAS E PREENCHIMENTOS
# ----------------------------------------------------------------------------

# Define os alinhamentos horizontais aceitos pelo openpyxl
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

# Define os alinhamentos verticais aceitos pelo openpyxl
ALINHAMENTOS_VERTICAIS = {
    "top",
    "center",
    "bottom",
    "justify",
    "distributed",
}

# Define os tipos de borda aceitos pelo openpyxl
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

# Define os tipos de preenchimento mais utilizados
TIPOS_PREENCHIMENTO = {
    "solid",
    "darkDown",
    "darkGray",
    "darkGrid",
    "darkHorizontal",
    "darkTrellis",
    "darkUp",
    "darkVertical",
    "gray0625",
    "gray125",
    "lightDown",
    "lightGray",
    "lightGrid",
    "lightHorizontal",
    "lightTrellis",
    "lightUp",
    "lightVertical",
    "mediumGray",
}

# Define o alinhamento horizontal padrão para cabeçalhos
ALINHAMENTO_CABECALHO_HORIZONTAL = "center"

# Define o alinhamento vertical padrão para cabeçalhos
ALINHAMENTO_CABECALHO_VERTICAL = "center"

# Define o tipo padrão de borda utilizado pela biblioteca
TIPO_BORDA_PADRAO = "thin"

# Define o tipo padrão de preenchimento utilizado pela biblioteca
TIPO_PREENCHIMENTO_PADRAO = "solid"


# ----------------------------------------------------------------------------
# CONFIGURAÇÕES PADRÃO
# ----------------------------------------------------------------------------

# Define o nome padrão da primeira planilha de um arquivo novo
NOME_PLANILHA_PADRAO = "Planilha1"

# Define a linha padrão usada como cabeçalho em tabelas e importações
LINHA_CABECALHO_PADRAO = 1

# Define a primeira linha padrão utilizada para dados
LINHA_DADOS_PADRAO = 2

# Define a margem padrão utilizada no ajuste automático de colunas
MARGEM_LARGURA_COLUNA = 2

# Define a largura mínima utilizada no ajuste automático
LARGURA_MINIMA_COLUNA = 1

# Define a largura máxima padrão utilizada no ajuste automático
LARGURA_MAXIMA_AJUSTE = 80

# Define o nome padrão aplicado às tabelas estruturadas
NOME_TABELA_PADRAO = "TabelaDados"

# Define o estilo padrão utilizado em tabelas estruturadas
ESTILO_TABELA_PADRAO = "TableStyleMedium2"

# Define o mecanismo padrão utilizado pelo pandas para arquivos modernos
ENGINE_EXCEL_PADRAO = "openpyxl"

# Define a orientação padrão das páginas impressas
ORIENTACAO_PAGINA_PADRAO = "portrait"

# Define o tamanho de papel padrão utilizado nas configurações de impressão
TAMANHO_PAPEL_PADRAO = "9"


# ----------------------------------------------------------------------------
# ENUMERAÇÕES
# ----------------------------------------------------------------------------

class TipoArquivoExcel(str, Enum):
    """Representa os principais tipos de arquivos Excel."""

    # Representa uma pasta de trabalho moderna sem macros
    WORKBOOK = ".xlsx"

    # Representa uma pasta de trabalho moderna com macros
    WORKBOOK_MACRO = ".xlsm"

    # Representa um modelo moderno sem macros
    MODELO = ".xltx"

    # Representa um modelo moderno com macros
    MODELO_MACRO = ".xltm"

    # Representa o formato legado do Excel
    LEGADO = ".xls"


class TipoAlinhamentoHorizontal(str, Enum):
    """Representa alinhamentos horizontais de células."""

    # Mantém o alinhamento geral definido pelo Excel
    GERAL = "general"

    # Alinha o conteúdo à esquerda
    ESQUERDA = "left"

    # Centraliza o conteúdo horizontalmente
    CENTRO = "center"

    # Alinha o conteúdo à direita
    DIREITA = "right"

    # Justifica o conteúdo horizontalmente
    JUSTIFICADO = "justify"

    # Distribui o conteúdo ao longo da célula
    DISTRIBUIDO = "distributed"


class TipoAlinhamentoVertical(str, Enum):
    """Representa alinhamentos verticais de células."""

    # Alinha o conteúdo ao topo da célula
    TOPO = "top"

    # Centraliza o conteúdo verticalmente
    CENTRO = "center"

    # Alinha o conteúdo à parte inferior da célula
    INFERIOR = "bottom"

    # Justifica o conteúdo verticalmente
    JUSTIFICADO = "justify"

    # Distribui o conteúdo verticalmente
    DISTRIBUIDO = "distributed"


class TipoBorda(str, Enum):
    """Representa estilos comuns de borda do Excel."""

    # Representa uma borda fina
    FINA = "thin"

    # Representa uma borda média
    MEDIA = "medium"

    # Representa uma borda grossa
    GROSSA = "thick"

    # Representa uma borda dupla
    DUPLA = "double"

    # Representa uma borda pontilhada
    PONTILHADA = "dotted"

    # Representa uma borda tracejada
    TRACEJADA = "dashed"


class TipoPreenchimento(str, Enum):
    """Representa tipos comuns de preenchimento de células."""

    # Representa um preenchimento sólido
    SOLIDO = "solid"

    # Representa um preenchimento cinza claro
    CINZA_CLARO = "lightGray"

    # Representa um preenchimento cinza médio
    CINZA_MEDIO = "mediumGray"

    # Representa um preenchimento cinza escuro
    CINZA_ESCURO = "darkGray"


class TipoFormatoCelula(str, Enum):
    """Representa categorias comuns de formatos de células."""

    # Representa um valor textual
    TEXTO = "texto"

    # Representa um número inteiro
    INTEIRO = "inteiro"

    # Representa um número decimal
    DECIMAL = "decimal"

    # Representa um valor monetário
    MOEDA = "moeda"

    # Representa um valor percentual
    PERCENTUAL = "percentual"

    # Representa uma data
    DATA = "data"

    # Representa uma hora
    HORA = "hora"

    # Representa uma data combinada com hora
    DATA_HORA = "data_hora"


class TipoOperacaoArquivo(str, Enum):
    """Representa operações comuns realizadas em arquivos Excel."""

    # Representa a criação de um novo arquivo
    CRIAR = "criar"

    # Representa a abertura de um arquivo existente
    ABRIR = "abrir"

    # Representa o salvamento do arquivo atual
    SALVAR = "salvar"

    # Representa uma cópia para outro caminho
    COPIAR = "copiar"

    # Representa a movimentação para outro caminho
    MOVER = "mover"

    # Representa a alteração do nome do arquivo
    RENOMEAR = "renomear"

    # Representa a exclusão do arquivo
    EXCLUIR = "excluir"

    # Representa a criação de uma cópia de segurança
    BACKUP = "backup"


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

# Define explicitamente os objetos públicos disponibilizados pelo módulo
__all__ = [
    "EXTENSOES_EXCEL",
    "EXTENSOES_OPENPYXL",
    "EXTENSOES_COM_MACRO",
    "EXTENSOES_MODELO",
    "EXTENSOES_WORKBOOK",
    "EXTENSAO_EXCEL_LEGADA",
    "EXTENSAO_EXCEL_PADRAO",
    "EXTENSAO_EXCEL_MACRO",
    "LIMITE_LINHAS_EXCEL",
    "LIMITE_COLUNAS_EXCEL",
    "ULTIMA_COLUNA_EXCEL",
    "LIMITE_CARACTERES_CELULA",
    "LIMITE_CARACTERES_PLANILHA",
    "LIMITE_CARACTERES_CABECALHO_RODAPE",
    "LIMITE_CARACTERES_NOME_DEFINIDO",
    "LARGURA_MAXIMA_COLUNA",
    "ALTURA_MAXIMA_LINHA",
    "LARGURA_PADRAO_COLUNA",
    "ALTURA_PADRAO_LINHA",
    "CARACTERES_INVALIDOS_PLANILHA",
    "CARACTERES_INVALIDOS_ARQUIVO",
    "NOMES_RESERVADOS_WINDOWS",
    "FORMATOS_DATA",
    "FORMATOS_HORA",
    "FORMATOS_DATA_HORA",
    "FORMATO_DATA_PADRAO",
    "FORMATO_DATA_HORA_PADRAO",
    "FORMATOS_NUMERO",
    "FORMATOS_MOEDA",
    "FORMATOS_PERCENTUAL",
    "FORMATOS_TEXTO",
    "FORMATO_NUMERO_PADRAO",
    "FORMATO_MOEDA_PADRAO",
    "FORMATO_PERCENTUAL_PADRAO",
    "CORES_EXCEL",
    "CORES_CORPORATIVAS",
    "CORES_TEXTO_MODELO",
    "COR_CABECALHO_PADRAO",
    "COR_TEXTO_CABECALHO_PADRAO",
    "ALINHAMENTOS_HORIZONTAIS",
    "ALINHAMENTOS_VERTICAIS",
    "TIPOS_BORDA",
    "TIPOS_PREENCHIMENTO",
    "ALINHAMENTO_CABECALHO_HORIZONTAL",
    "ALINHAMENTO_CABECALHO_VERTICAL",
    "TIPO_BORDA_PADRAO",
    "TIPO_PREENCHIMENTO_PADRAO",
    "NOME_PLANILHA_PADRAO",
    "LINHA_CABECALHO_PADRAO",
    "LINHA_DADOS_PADRAO",
    "MARGEM_LARGURA_COLUNA",
    "LARGURA_MINIMA_COLUNA",
    "LARGURA_MAXIMA_AJUSTE",
    "NOME_TABELA_PADRAO",
    "ESTILO_TABELA_PADRAO",
    "ENGINE_EXCEL_PADRAO",
    "ORIENTACAO_PAGINA_PADRAO",
    "TAMANHO_PAPEL_PADRAO",
    "TipoArquivoExcel",
    "TipoAlinhamentoHorizontal",
    "TipoAlinhamentoVertical",
    "TipoBorda",
    "TipoPreenchimento",
    "TipoFormatoCelula",
    "TipoOperacaoArquivo",
]
