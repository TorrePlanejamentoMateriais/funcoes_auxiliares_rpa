"""
============================================================
Módulo: EXCEL / arquivos.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versao: 1.0.0
Descrição:
    Funções reutilizáveis para criar, abrir, salvar, copiar, mover,
    renomear, listar e validar arquivos Excel. O modulo utiliza openpyxl
    para arquivos .xlsx, .xlsm, .xltx e .xltm.
Dependencias:
    - openpyxl
    - pathlib
    - shutil
Observações:
    - O formato legado .xls não é suportado pelo openpyxl.
    - Para preservar macros em .xlsm e .xltm, use manter_vba=True.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------
# Habilita a avaliação adiada das anotações de tipo
from __future__ import annotations

# Importa o módulo utilizado para copiar e mover arquivos
import shutil
# Importa Path para manipulação segura e multiplataforma de caminhos
from pathlib import Path
# Importa tipos auxiliares utilizados nas anotações das funções
from typing import Any, Iterable

# Importa os recursos principais para criar e abrir arquivos Excel
from openpyxl import Workbook, load_workbook
# Importa a classe concreta para validar objetos Workbook
from openpyxl.workbook.workbook import Workbook as OpenpyxlWorkbook


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------
# Define as extensões que podem ser manipuladas pelo openpyxl
EXTENSOES_EXCEL_OPENPYXL = {".xlsx", ".xlsm", ".xltx", ".xltm"}
# Define as extensões que podem armazenar projetos VBA
EXTENSOES_EXCEL_COM_MACRO = {".xlsm", ".xltm"}
# Define as extensões utilizadas por modelos do Excel
EXTENSOES_EXCEL_MODELO = {".xltx", ".xltm"}
# Define a extensão legada que não é aberta pelo openpyxl
EXTENSAO_EXCEL_LEGADA = ".xls"


# ----------------------------------------------------------------------------
# EXCEÇÕES
# ----------------------------------------------------------------------------
class ExcelArquivosError(Exception):
    """Erro base para operacoes com arquivos Excel."""


class ArquivoExcelInvalidoError(ExcelArquivosError):
    """O caminho ou a extensao do arquivo Excel e invalido."""


class ArquivoExcelNaoEncontradoError(ExcelArquivosError, FileNotFoundError):
    """O arquivo Excel solicitado nao foi encontrado."""


class ArquivoExcelExistenteError(ExcelArquivosError, FileExistsError):
    """O arquivo de destino ja existe e a sobrescrita nao foi autorizada."""


class WorkbookAberturaError(ExcelArquivosError):
    """O workbook nao pode ser aberto pelo openpyxl."""


class WorkbookSalvamentoError(ExcelArquivosError):
    """O workbook nao pode ser salvo no caminho solicitado."""


class OperacaoArquivoExcelError(ExcelArquivosError):
    """Uma operacao de copia, movimentacao ou exclusao falhou."""


# ----------------------------------------------------------------------------
# FUNÇÕES INTERNAS DE VALIDAÇÃO
# ----------------------------------------------------------------------------
def _validar_booleano(valor: bool, nome: str) -> bool:
    """Valida estritamente um parametro booleano."""
    # Verifica se not isinstance(valor, bool)
    if not isinstance(valor, bool):
        # Interrompe a execução e informa o erro encontrado
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")
    # Retorna o resultado processado pela função
    return valor


def _validar_texto(
    valor: str,
    nome: str,
    permitir_vazio: bool = False,
) -> str:
    """Valida e normaliza um parametro textual."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(permitir_vazio, "permitir_vazio")
    # Verifica se not isinstance(valor, str)
    if not isinstance(valor, str):
        # Interrompe a execução e informa o erro encontrado
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")
    # Armazena o valor processado em texto
    texto = valor.strip()
    # Verifica se not texto and not permitir_vazio
    if not texto and not permitir_vazio:
        # Interrompe a execução e informa o erro encontrado
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")
    # Retorna o resultado processado pela função
    return texto


def _normalizar_extensao(extensao: str) -> str:
    """Normaliza uma extensao para o formato .xlsx."""
    # Armazena o valor processado em texto
    texto = _validar_texto(extensao, "extensao").lower()
    # Retorna o resultado processado pela função
    return texto if texto.startswith(".") else f".{texto}"


def _normalizar_extensoes(
    extensoes: str | Iterable[str] | None,
) -> set[str]:
    """Normaliza uma ou varias extensoes permitidas."""
    # Verifica se extensoes is None
    if extensoes is None:
        # Retorna o resultado processado pela função
        return set(EXTENSOES_EXCEL_OPENPYXL)
    # Verifica se isinstance(extensoes, str)
    if isinstance(extensoes, str):
        # Armazena o valor processado em extensoes
        extensoes = [extensoes]
    # Verifica se not isinstance(extensoes, Iterable)
    if not isinstance(extensoes, Iterable):
        # Interrompe a execução e informa o erro encontrado
        raise TypeError("O parametro 'extensoes' deve ser texto, iteravel ou None.")
    # Armazena o valor processado em resultado
    resultado = {_normalizar_extensao(extensao) for extensao in extensoes}
    # Verifica se not resultado
    if not resultado:
        # Interrompe a execução e informa o erro encontrado
        raise ValueError("Informe ao menos uma extensao.")
    # Retorna o resultado processado pela função
    return resultado


def _validar_workbook(workbook: Any) -> OpenpyxlWorkbook:
    """Valida se o objeto informado e um Workbook do openpyxl."""
    # Verifica se not isinstance(workbook, OpenpyxlWorkbook)
    if not isinstance(workbook, OpenpyxlWorkbook):
        # Interrompe a execução e informa o erro encontrado
        raise TypeError("O parametro 'workbook' deve ser um Workbook do openpyxl.")
    # Retorna o resultado processado pela função
    return workbook


# ----------------------------------------------------------------------------
# FUNÇÕES PÚBLICAS
# ----------------------------------------------------------------------------
def normalizar_caminho(caminho: str | Path) -> Path:
    """Converte string ou Path em caminho absoluto normalizado."""
    # Verifica se not isinstance(caminho, (str, Path))
    if not isinstance(caminho, (str, Path)):
        # Interrompe a execução e informa o erro encontrado
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    # Verifica se isinstance(caminho, str) and not caminho.strip()
    if isinstance(caminho, str) and not caminho.strip():
        # Interrompe a execução e informa o erro encontrado
        raise ValueError("O parametro 'caminho' nao pode estar vazio.")
    # Retorna o resultado processado pela função
    return Path(caminho).expanduser().resolve()


def obter_extensao_excel(caminho: str | Path) -> str:
    """Retorna a extensao do arquivo em letras minusculas."""
    # Retorna o resultado processado pela função
    return normalizar_caminho(caminho).suffix.lower()


def extensao_excel_valida(
    caminho_ou_extensao: str | Path,
    permitir_xls: bool = False,
) -> bool:
    """Verifica se uma extensao pertence a um formato Excel conhecido."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(permitir_xls, "permitir_xls")
    # Verifica se isinstance(caminho_ou_extensao, Path)
    if isinstance(caminho_ou_extensao, Path):
        # Armazena o valor processado em extensao
        extensao = caminho_ou_extensao.suffix.lower()
    # Verifica se isinstance(caminho_ou_extensao, str)
    elif isinstance(caminho_ou_extensao, str):
        # Armazena o valor processado em texto
        texto = caminho_ou_extensao.strip()
        # Verifica se not texto
        if not texto:
            # Retorna o resultado processado pela função
            return False
        # Armazena o valor processado em extensao
        extensao = (
            texto.lower()
            if texto.startswith(".") and "/" not in texto and "\\" not in texto
            else Path(texto).suffix.lower()
        )
    else:
        # Interrompe a execução e informa o erro encontrado
        raise TypeError("O parametro deve ser string ou Path.")
    # Armazena o valor processado em permitidas
    permitidas = set(EXTENSOES_EXCEL_OPENPYXL)
    # Verifica se permitir_xls
    if permitir_xls:
        # Executa a operação necessária nesta etapa
        permitidas.add(EXTENSAO_EXCEL_LEGADA)
    # Retorna o resultado processado pela função
    return extensao in permitidas


def validar_caminho_excel(
    caminho: str | Path,
    deve_existir: bool = True,
    extensoes: str | Iterable[str] | None = None,
) -> Path:
    """Valida caminho, existencia e extensao de um arquivo Excel."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(deve_existir, "deve_existir")
    # Armazena o valor processado em arquivo
    arquivo = normalizar_caminho(caminho)
    # Armazena o valor processado em permitidas
    permitidas = _normalizar_extensoes(extensoes)
    # Verifica se arquivo.suffix.lower() not in permitidas
    if arquivo.suffix.lower() not in permitidas:
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelInvalidoError(
            f"Extensao '{arquivo.suffix}' nao suportada. "
            f"Permitidas: {sorted(permitidas)}."
        )
    # Verifica se deve_existir and not arquivo.is_file()
    if deve_existir and not arquivo.is_file():
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelNaoEncontradoError(
            f"O arquivo Excel nao existe: {arquivo}"
        )
    # Retorna o resultado processado pela função
    return arquivo


def arquivo_excel_valido(
    caminho: str | Path,
    verificar_abertura: bool = False,
    manter_vba: bool | None = None,
) -> bool:
    """Verifica caminho, extensao e opcionalmente a abertura do workbook."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(verificar_abertura, "verificar_abertura")
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Armazena o valor processado em arquivo
        arquivo = validar_caminho_excel(caminho, deve_existir=True)
        # Verifica se not verificar_abertura
        if not verificar_abertura:
            # Retorna o resultado processado pela função
            return True
        # Armazena o valor processado em preservar_vba
        preservar_vba = (
            arquivo.suffix.lower() in EXTENSOES_EXCEL_COM_MACRO
            if manter_vba is None
            else _validar_booleano(manter_vba, "manter_vba")
        )
        # Armazena o valor processado em workbook
        workbook = load_workbook(
            arquivo,
            read_only=True,
            data_only=False,
            keep_vba=preservar_vba,
        )
        # Fecha o workbook para liberar os recursos utilizados
        workbook.close()
        # Retorna o resultado processado pela função
        return True
    except Exception:
        # Converte a falha original em uma exceção específica do módulo
        return False


def criar_workbook(
    titulo_planilha: str = "Planilha1",
    remover_planilha_padrao: bool = False,
) -> OpenpyxlWorkbook:
    """Cria um novo Workbook com uma planilha inicial configuravel."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(remover_planilha_padrao, "remover_planilha_padrao")
    # Armazena o valor processado em titulo
    titulo = _validar_texto(titulo_planilha, "titulo_planilha")
    # Verifica se len(titulo) > 31
    if len(titulo) > 31:
        # Interrompe a execução e informa o erro encontrado
        raise ValueError("O titulo da planilha nao pode ultrapassar 31 caracteres.")
    # Armazena o valor processado em caracteres_invalidos
    caracteres_invalidos = set("[]:*?/\\")
    # Verifica se any(caractere in titulo for caractere in caracteres_invalidos)
    if any(caractere in titulo for caractere in caracteres_invalidos):
        # Interrompe a execução e informa o erro encontrado
        raise ValueError("O titulo da planilha possui caracteres invalidos.")
    # Armazena o valor processado em workbook
    workbook = Workbook()
    # Armazena o valor processado em planilha
    planilha = workbook.active
    # Armazena o valor processado em planilha.title
    planilha.title = titulo
    # Verifica se remover_planilha_padrao
    if remover_planilha_padrao:
        # Remove a planilha inicial criada automaticamente
        workbook.remove(planilha)
    # Retorna o resultado processado pela função
    return workbook


def abrir_workbook(
    caminho: str | Path,
    somente_leitura: bool = False,
    somente_valores: bool = False,
    manter_vba: bool | None = None,
    manter_links: bool = True,
    rico_texto: bool = False,
) -> OpenpyxlWorkbook:
    """Abre um arquivo Excel utilizando openpyxl.load_workbook."""
    # Percorre os elementos da coleção para processar cada item
    for nome, valor in {
        "somente_leitura": somente_leitura,
        "somente_valores": somente_valores,
        "manter_links": manter_links,
        "rico_texto": rico_texto,
    }.items():
        # Executa a operação necessária nesta etapa
        _validar_booleano(valor, nome)
    # Armazena o valor processado em arquivo
    arquivo = validar_caminho_excel(caminho, deve_existir=True)
    # Armazena o valor processado em preservar_vba
    preservar_vba = (
        arquivo.suffix.lower() in EXTENSOES_EXCEL_COM_MACRO
        if manter_vba is None
        else _validar_booleano(manter_vba, "manter_vba")
    )
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Retorna o resultado processado pela função
        return load_workbook(
            filename=arquivo,
            read_only=somente_leitura,
            data_only=somente_valores,
            keep_vba=preservar_vba,
            keep_links=manter_links,
            rich_text=rico_texto,
        )
    except Exception as error:
        # Converte a falha original em uma exceção específica do módulo
        raise WorkbookAberturaError(
            f"Nao foi possivel abrir o workbook: {arquivo}"
        ) from error


def salvar_workbook(
    workbook: OpenpyxlWorkbook,
    caminho: str | Path,
    sobrescrever: bool = True,
    criar_diretorios: bool = True,
) -> Path:
    """Salva um Workbook no caminho informado e retorna o Path final."""
    # Executa a operação necessária nesta etapa
    _validar_workbook(workbook)
    # Executa a operação necessária nesta etapa
    _validar_booleano(sobrescrever, "sobrescrever")
    # Executa a operação necessária nesta etapa
    _validar_booleano(criar_diretorios, "criar_diretorios")
    # Armazena o valor processado em destino
    destino = validar_caminho_excel(caminho, deve_existir=False)
    # Verifica se destino.exists() and not sobrescrever
    if destino.exists() and not sobrescrever:
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelExistenteError(f"O arquivo ja existe: {destino}")
    # Verifica se criar_diretorios
    if criar_diretorios:
        # Cria o diretório de destino quando necessário
        destino.parent.mkdir(parents=True, exist_ok=True)
    # Verifica se not destino.parent.is_dir()
    elif not destino.parent.is_dir():
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelNaoEncontradoError(
            f"O diretorio de destino nao existe: {destino.parent}"
        )
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Executa a operação necessária nesta etapa
        workbook.save(destino)
        # Retorna o resultado processado pela função
        return destino
    except Exception as error:
        # Converte a falha original em uma exceção específica do módulo
        raise WorkbookSalvamentoError(
            f"Nao foi possivel salvar o workbook: {destino}"
        ) from error


def salvar_como(
    workbook: OpenpyxlWorkbook,
    caminho: str | Path,
    sobrescrever: bool = False,
) -> Path:
    """Salva o Workbook em outro caminho, sem sobrescrever por padrao."""
    # Retorna o resultado processado pela função
    return salvar_workbook(
        workbook,
        caminho,
        sobrescrever=sobrescrever,
        criar_diretorios=True,
    )


def fechar_workbook(workbook: OpenpyxlWorkbook) -> None:
    """Fecha um Workbook, liberando os recursos associados."""
    # Executa a operação necessária nesta etapa
    _validar_workbook(workbook)
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Fecha o workbook para liberar os recursos utilizados
        workbook.close()
    except Exception as error:
        # Converte a falha original em uma exceção específica do módulo
        raise ExcelArquivosError("Nao foi possivel fechar o workbook.") from error


def criar_arquivo_excel(
    caminho: str | Path,
    titulo_planilha: str = "Planilha1",
    sobrescrever: bool = False,
) -> Path:
    """Cria e salva um novo arquivo Excel em uma unica operacao."""
    # Armazena o valor processado em workbook
    workbook = criar_workbook(titulo_planilha=titulo_planilha)
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Retorna o resultado processado pela função
        return salvar_workbook(
            workbook,
            caminho,
            sobrescrever=sobrescrever,
            criar_diretorios=True,
        )
    finally:
        # Fecha o workbook para liberar os recursos utilizados
        workbook.close()


def copiar_arquivo_excel(
    origem: str | Path,
    destino: str | Path,
    sobrescrever: bool = False,
    criar_diretorios: bool = True,
) -> Path:
    """Copia um arquivo Excel preservando seus metadados."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(sobrescrever, "sobrescrever")
    # Executa a operação necessária nesta etapa
    _validar_booleano(criar_diretorios, "criar_diretorios")
    # Armazena o valor processado em arquivo_origem
    arquivo_origem = validar_caminho_excel(origem, deve_existir=True)
    # Armazena o valor processado em arquivo_destino
    arquivo_destino = validar_caminho_excel(destino, deve_existir=False)
    # Verifica se arquivo_destino.exists() and not sobrescrever
    if arquivo_destino.exists() and not sobrescrever:
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelExistenteError(
            f"O arquivo de destino ja existe: {arquivo_destino}"
        )
    # Verifica se criar_diretorios
    if criar_diretorios:
        # Cria o diretório de destino quando necessário
        arquivo_destino.parent.mkdir(parents=True, exist_ok=True)
    # Verifica se not arquivo_destino.parent.is_dir()
    elif not arquivo_destino.parent.is_dir():
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelNaoEncontradoError(
            f"O diretorio de destino nao existe: {arquivo_destino.parent}"
        )
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Copia o arquivo preservando os metadados disponíveis
        shutil.copy2(arquivo_origem, arquivo_destino)
        # Retorna o resultado processado pela função
        return arquivo_destino
    except OSError as error:
        # Converte a falha original em uma exceção específica do módulo
        raise OperacaoArquivoExcelError(
            f"Nao foi possivel copiar '{arquivo_origem}' para '{arquivo_destino}'."
        ) from error


def mover_arquivo_excel(
    origem: str | Path,
    destino: str | Path,
    sobrescrever: bool = False,
    criar_diretorios: bool = True,
) -> Path:
    """Move um arquivo Excel para outro caminho."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(sobrescrever, "sobrescrever")
    # Executa a operação necessária nesta etapa
    _validar_booleano(criar_diretorios, "criar_diretorios")
    # Armazena o valor processado em arquivo_origem
    arquivo_origem = validar_caminho_excel(origem, deve_existir=True)
    # Armazena o valor processado em arquivo_destino
    arquivo_destino = validar_caminho_excel(destino, deve_existir=False)
    # Verifica se arquivo_destino.exists()
    if arquivo_destino.exists():
        # Verifica se not sobrescrever
        if not sobrescrever:
            # Interrompe a execução e informa o erro encontrado
            raise ArquivoExcelExistenteError(
                f"O arquivo de destino ja existe: {arquivo_destino}"
            )
        # Remove o arquivo existente antes da movimentação
        arquivo_destino.unlink()
    # Verifica se criar_diretorios
    if criar_diretorios:
        # Cria o diretório de destino quando necessário
        arquivo_destino.parent.mkdir(parents=True, exist_ok=True)
    # Verifica se not arquivo_destino.parent.is_dir()
    elif not arquivo_destino.parent.is_dir():
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelNaoEncontradoError(
            f"O diretorio de destino nao existe: {arquivo_destino.parent}"
        )
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Retorna o resultado processado pela função
        return Path(shutil.move(str(arquivo_origem), str(arquivo_destino))).resolve()
    except OSError as error:
        # Converte a falha original em uma exceção específica do módulo
        raise OperacaoArquivoExcelError(
            f"Nao foi possivel mover '{arquivo_origem}' para '{arquivo_destino}'."
        ) from error


def renomear_arquivo_excel(
    caminho: str | Path,
    novo_nome: str,
    sobrescrever: bool = False,
) -> Path:
    """Renomeia um arquivo Excel dentro do diretorio atual."""
    # Armazena o valor processado em origem
    origem = validar_caminho_excel(caminho, deve_existir=True)
    # Armazena o valor processado em nome
    nome = _validar_texto(novo_nome, "novo_nome")
    # Armazena o valor processado em destino
    destino = origem.with_name(nome)
    # Executa a operação necessária nesta etapa
    validar_caminho_excel(destino, deve_existir=False)
    # Retorna o resultado processado pela função
    return mover_arquivo_excel(
        origem,
        destino,
        sobrescrever=sobrescrever,
        criar_diretorios=False,
    )


def excluir_arquivo_excel(
    caminho: str | Path,
    ignorar_inexistente: bool = False,
) -> bool:
    """Exclui um arquivo Excel e retorna True quando removido."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(ignorar_inexistente, "ignorar_inexistente")
    # Armazena o valor processado em arquivo
    arquivo = normalizar_caminho(caminho)
    # Verifica se not arquivo.exists()
    if not arquivo.exists():
        # Verifica se ignorar_inexistente
        if ignorar_inexistente:
            # Retorna o resultado processado pela função
            return False
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelNaoEncontradoError(
            f"O arquivo Excel nao existe: {arquivo}"
        )
    # Verifica se not arquivo.is_file() or not extensao_excel_valida(arquivo)
    if not arquivo.is_file() or not extensao_excel_valida(arquivo):
        # Interrompe a execução e informa o erro encontrado
        raise ArquivoExcelInvalidoError(f"O caminho nao e um arquivo Excel: {arquivo}")
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Remove o arquivo Excel do sistema de arquivos
        arquivo.unlink()
        # Retorna o resultado processado pela função
        return True
    except OSError as error:
        # Converte a falha original em uma exceção específica do módulo
        raise OperacaoArquivoExcelError(
            f"Nao foi possivel excluir o arquivo: {arquivo}"
        ) from error


def listar_arquivos_excel(
    diretorio: str | Path,
    recursivo: bool = False,
    extensoes: str | Iterable[str] | None = None,
    ordenar: bool = True,
) -> list[Path]:
    """Lista arquivos Excel de um diretorio."""
    # Executa a operação necessária nesta etapa
    _validar_booleano(recursivo, "recursivo")
    # Executa a operação necessária nesta etapa
    _validar_booleano(ordenar, "ordenar")
    # Armazena o valor processado em pasta
    pasta = normalizar_caminho(diretorio)
    # Verifica se not pasta.is_dir()
    if not pasta.is_dir():
        # Interrompe a execução e informa o erro encontrado
        raise NotADirectoryError(f"O diretorio nao existe: {pasta}")
    # Armazena o valor processado em permitidas
    permitidas = _normalizar_extensoes(extensoes)
    # Armazena o valor processado em iterador
    iterador = pasta.rglob("*") if recursivo else pasta.glob("*")
    # Armazena o valor processado em arquivos
    arquivos = [
        caminho.resolve()
        for caminho in iterador
        if caminho.is_file() and caminho.suffix.lower() in permitidas
    ]
    # Verifica se ordenar
    if ordenar:
        # Executa a operação necessária nesta etapa
        arquivos.sort(key=lambda caminho: str(caminho).casefold())
    # Retorna o resultado processado pela função
    return arquivos


def obter_tamanho_arquivo(caminho: str | Path) -> int:
    """Retorna o tamanho do arquivo Excel em bytes."""
    # Armazena o valor processado em arquivo
    arquivo = validar_caminho_excel(caminho, deve_existir=True)
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Retorna o resultado processado pela função
        return arquivo.stat().st_size
    except OSError as error:
        # Converte a falha original em uma exceção específica do módulo
        raise OperacaoArquivoExcelError(
            f"Nao foi possivel consultar o tamanho de: {arquivo}"
        ) from error


def obter_dados_arquivo_excel(caminho: str | Path) -> dict[str, Any]:
    """Retorna metadados principais de um arquivo Excel."""
    # Armazena o valor processado em arquivo
    arquivo = validar_caminho_excel(caminho, deve_existir=True)
    # Executa a operação principal com tratamento de possíveis falhas
    try:
        # Armazena o valor processado em estatisticas
        estatisticas = arquivo.stat()
    except OSError as error:
        # Converte a falha original em uma exceção específica do módulo
        raise OperacaoArquivoExcelError(
            f"Nao foi possivel consultar os dados do arquivo: {arquivo}"
        ) from error
    # Armazena o valor processado em extensao
    extensao = arquivo.suffix.lower()
    # Retorna o resultado processado pela função
    return {
        "nome": arquivo.name,
        "stem": arquivo.stem,
        "extensao": extensao,
        "caminho": arquivo,
        "diretorio": arquivo.parent,
        "tamanho_bytes": estatisticas.st_size,
        "modificado_em": estatisticas.st_mtime,
        "possui_macro": extensao in EXTENSOES_EXCEL_COM_MACRO,
        "e_modelo": extensao in EXTENSOES_EXCEL_MODELO,
    }


def caminho_excel_unico(caminho: str | Path) -> Path:
    """Gera um caminho Excel livre adicionando sufixo numerico."""
    # Armazena o valor processado em original
    original = validar_caminho_excel(caminho, deve_existir=False)
    # Verifica se not original.exists()
    if not original.exists():
        # Retorna o resultado processado pela função
        return original
    # Armazena o valor processado em contador
    contador = 1
    # Repete a operação até localizar um resultado disponível
    while True:
        # Armazena o valor processado em candidato
        candidato = original.with_name(
            f"{original.stem}_{contador}{original.suffix}"
        )
        # Verifica se not candidato.exists()
        if not candidato.exists():
            # Retorna o resultado processado pela função
            return candidato
        # Atualiza o contador utilizado para gerar um nome disponível
        contador += 1


def criar_backup_excel(
    caminho: str | Path,
    diretorio_backup: str | Path | None = None,
    sufixo: str = "_backup",
    sobrescrever: bool = False,
) -> Path:
    """Cria uma copia de seguranca de um arquivo Excel."""
    # Armazena o valor processado em origem
    origem = validar_caminho_excel(caminho, deve_existir=True)
    # Armazena o valor processado em sufixo_validado
    sufixo_validado = _validar_texto(sufixo, "sufixo")
    # Armazena o valor processado em pasta_destino
    pasta_destino = (
        origem.parent
        if diretorio_backup is None
        else normalizar_caminho(diretorio_backup)
    )
    # Armazena o valor processado em destino
    destino = pasta_destino / f"{origem.stem}{sufixo_validado}{origem.suffix}"
    # Retorna o resultado processado pela função
    return copiar_arquivo_excel(
        origem,
        destino,
        sobrescrever=sobrescrever,
        criar_diretorios=True,
    )


__all__ = [
    "EXTENSOES_EXCEL_OPENPYXL",
    "EXTENSOES_EXCEL_COM_MACRO",
    "EXTENSOES_EXCEL_MODELO",
    "EXTENSAO_EXCEL_LEGADA",
    "ExcelArquivosError",
    "ArquivoExcelInvalidoError",
    "ArquivoExcelNaoEncontradoError",
    "ArquivoExcelExistenteError",
    "WorkbookAberturaError",
    "WorkbookSalvamentoError",
    "OperacaoArquivoExcelError",
    "normalizar_caminho",
    "obter_extensao_excel",
    "extensao_excel_valida",
    "validar_caminho_excel",
    "arquivo_excel_valido",
    "criar_workbook",
    "abrir_workbook",
    "salvar_workbook",
    "salvar_como",
    "fechar_workbook",
    "criar_arquivo_excel",
    "copiar_arquivo_excel",
    "mover_arquivo_excel",
    "renomear_arquivo_excel",
    "excluir_arquivo_excel",
    "listar_arquivos_excel",
    "obter_tamanho_arquivo",
    "obter_dados_arquivo_excel",
    "caminho_excel_unico",
    "criar_backup_excel",
]
