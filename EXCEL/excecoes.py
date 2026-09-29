"""
============================================================
Módulo: EXCEL / excecoes.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criação: 28/09/2026
Última Alteração: 28/09/2026
Versão: 1.0.0
Descrição:
    Centraliza as exceções personalizadas utilizadas pelos módulos da
    biblioteca de automação de arquivos Excel. A hierarquia permite
    capturar erros gerais ou tratar falhas específicas por assunto.
Exceções disponíveis:
    - ExcelAutomacaoError
    - ExcelValidacaoError
    - ArquivoExcelError
    - ArquivoExcelInvalidoError
    - ArquivoExcelNaoEncontradoError
    - ArquivoExcelExistenteError
    - WorkbookError
    - WorkbookAberturaError
    - WorkbookSalvamentoError
    - WorkbookFechamentoError
    - PlanilhaError
    - PlanilhaInvalidaError
    - PlanilhaNaoEncontradaError
    - PlanilhaExistenteError
    - NomePlanilhaInvalidoError
    - CelulaError
    - CelulaInvalidaError
    - IntervaloError
    - IntervaloInvalidoError
    - ColunaError
    - ColunaInvalidaError
    - ColunaNaoEncontradaError
    - LinhaError
    - LinhaInvalidaError
    - TabelaExcelError
    - TabelaNaoEncontradaError
    - TabelaExistenteError
    - FormulaExcelError
    - FormulaInvalidaError
    - FormatacaoExcelError
    - ImportacaoExcelError
    - ExportacaoExcelError
    - DadosExcelError
    - DadosObrigatoriosError
    - OperacaoExcelError
Dependências:
    - Nenhuma dependência externa
Histórico:
    v1.0.0 - 28/09/2026
        - Criação inicial do módulo.
        - Inclusão da hierarquia central de exceções.
        - Inclusão de suporte a contexto e erro original.
        - Inclusão de funções auxiliares para tratamento padronizado.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTAÇÕES
# ----------------------------------------------------------------------------

# Importa Any para permitir informações adicionais de diferentes tipos
from typing import Any


# ----------------------------------------------------------------------------
# EXCEÇÃO BASE
# ----------------------------------------------------------------------------

class ExcelAutomacaoError(Exception):
    """
    Representa a exceção base da biblioteca de automação Excel.

    Todas as exceções específicas do módulo herdam desta classe. Isso
    permite capturar qualquer falha da biblioteca utilizando somente
    ExcelAutomacaoError, sem perder a possibilidade de tratamentos mais
    específicos.

    Args:
        mensagem (str):
            Descrição principal do erro ocorrido.
        contexto (dict[str, Any] | None):
            Informações adicionais relacionadas à operação.
        erro_original (Exception | None):
            Exceção original que causou a falha atual.

    Attributes:
        mensagem (str):
            Mensagem normalizada do erro.
        contexto (dict[str, Any]):
            Dicionário com informações adicionais.
        erro_original (Exception | None):
            Exceção original associada.
    """

    def __init__(
        self,
        mensagem: str,
        contexto: dict[str, Any] | None = None,
        erro_original: Exception | None = None,
    ) -> None:
        """
        Inicializa a exceção com mensagem, contexto e causa original.

        Raises:
            TypeError:
                Caso a mensagem não seja uma string.
            ValueError:
                Caso a mensagem esteja vazia.
            TypeError:
                Caso o contexto não seja um dicionário ou None.
            TypeError:
                Caso o erro original não seja uma exceção ou None.
        """

        # Verifica se a mensagem foi informada como string
        if not isinstance(mensagem, str):
            # Interrompe a execução quando a mensagem possui tipo inválido
            raise TypeError(
                "O parâmetro 'mensagem' deve ser uma string."
            )

        # Remove os espaços externos da mensagem
        mensagem_normalizada = mensagem.strip()

        # Verifica se a mensagem possui conteúdo
        if not mensagem_normalizada:
            # Interrompe a execução quando a mensagem está vazia
            raise ValueError(
                "O parâmetro 'mensagem' não pode estar vazio."
            )

        # Verifica se o contexto possui um tipo válido
        if contexto is not None and not isinstance(contexto, dict):
            # Interrompe a execução quando o contexto não é um dicionário
            raise TypeError(
                "O parâmetro 'contexto' deve ser um dicionário ou None."
            )

        # Verifica se o erro original é uma exceção válida
        if erro_original is not None and not isinstance(
            erro_original,
            Exception,
        ):
            # Interrompe a execução quando a causa possui tipo inválido
            raise TypeError(
                "O parâmetro 'erro_original' deve ser uma exceção ou None."
            )

        # Armazena a mensagem normalizada
        self.mensagem = mensagem_normalizada

        # Cria uma cópia do contexto para evitar alterações externas
        self.contexto = dict(contexto or {})

        # Armazena a exceção original associada
        self.erro_original = erro_original

        # Inicializa a classe Exception com a mensagem formatada
        super().__init__(self.formatar_mensagem())

    def formatar_mensagem(self) -> str:
        """
        Formata a mensagem da exceção incluindo o contexto disponível.

        Returns:
            str:
                Mensagem principal seguida pelo contexto, quando existente.
        """

        # Retorna somente a mensagem quando não existe contexto
        if not self.contexto:
            return self.mensagem

        # Converte cada informação de contexto para uma representação textual
        detalhes = ", ".join(
            f"{chave}={valor!r}"
            for chave, valor in self.contexto.items()
        )

        # Retorna a mensagem combinada com os detalhes da operação
        return f"{self.mensagem} Contexto: {detalhes}."

    def para_dicionario(self) -> dict[str, Any]:
        """
        Converte a exceção para um dicionário adequado para logs.

        Returns:
            dict[str, Any]:
                Dicionário contendo tipo, mensagem, contexto e causa.
        """

        # Cria a estrutura padronizada de informações do erro
        dados = {
            "tipo": self.__class__.__name__,
            "mensagem": self.mensagem,
            "contexto": dict(self.contexto),
            "erro_original": (
                None
                if self.erro_original is None
                else str(self.erro_original)
            ),
        }

        # Retorna os dados preparados para registro
        return dados


# ----------------------------------------------------------------------------
# EXCEÇÕES DE VALIDAÇÃO
# ----------------------------------------------------------------------------

class ExcelValidacaoError(ExcelAutomacaoError, ValueError):
    """Representa falhas de validação de parâmetros ou estruturas Excel."""


class DadosExcelError(ExcelValidacaoError):
    """Representa dados inválidos, inconsistentes ou incompatíveis."""


class DadosObrigatoriosError(DadosExcelError):
    """Indica a ausência de dados obrigatórios para uma operação."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE ARQUIVOS
# ----------------------------------------------------------------------------

class ArquivoExcelError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a arquivos Excel."""


class ArquivoExcelInvalidoError(ArquivoExcelError, ValueError):
    """Indica que um caminho, extensão ou conteúdo Excel é inválido."""


class ArquivoExcelNaoEncontradoError(
    ArquivoExcelError,
    FileNotFoundError,
):
    """Indica que o arquivo Excel solicitado não foi encontrado."""


class ArquivoExcelExistenteError(ArquivoExcelError, FileExistsError):
    """Indica que o arquivo de destino já existe."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE WORKBOOK
# ----------------------------------------------------------------------------

class WorkbookError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas ao Workbook."""


class WorkbookAberturaError(WorkbookError):
    """Indica que um Workbook não pôde ser aberto."""


class WorkbookSalvamentoError(WorkbookError):
    """Indica que um Workbook não pôde ser salvo."""


class WorkbookFechamentoError(WorkbookError):
    """Indica que um Workbook não pôde ser fechado corretamente."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE PLANILHAS
# ----------------------------------------------------------------------------

class PlanilhaError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a planilhas."""


class PlanilhaInvalidaError(PlanilhaError, ValueError):
    """Indica que o objeto ou o nome da planilha é inválido."""


class PlanilhaNaoEncontradaError(PlanilhaError, KeyError):
    """Indica que a planilha solicitada não existe no Workbook."""


class PlanilhaExistenteError(PlanilhaError):
    """Indica que já existe uma planilha com o nome informado."""


class NomePlanilhaInvalidoError(PlanilhaInvalidaError):
    """Indica que o nome da planilha viola uma regra do Excel."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE CÉLULAS E INTERVALOS
# ----------------------------------------------------------------------------

class CelulaError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a células."""


class CelulaInvalidaError(CelulaError, ValueError):
    """Indica que uma referência ou um valor de célula é inválido."""


class IntervaloError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a intervalos de células."""


class IntervaloInvalidoError(IntervaloError, ValueError):
    """Indica que a referência ou os limites do intervalo são inválidos."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE COLUNAS E LINHAS
# ----------------------------------------------------------------------------

class ColunaError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a colunas."""


class ColunaInvalidaError(ColunaError, ValueError):
    """Indica que uma referência ou configuração de coluna é inválida."""


class ColunaNaoEncontradaError(ColunaError, KeyError):
    """Indica que a coluna ou o cabeçalho solicitado não foi encontrado."""


class LinhaError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a linhas."""


class LinhaInvalidaError(LinhaError, ValueError):
    """Indica que uma referência ou configuração de linha é inválida."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE TABELAS, FÓRMULAS E FORMATAÇÃO
# ----------------------------------------------------------------------------

class TabelaExcelError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a tabelas estruturadas."""


class TabelaNaoEncontradaError(TabelaExcelError, KeyError):
    """Indica que a tabela estruturada solicitada não foi encontrada."""


class TabelaExistenteError(TabelaExcelError):
    """Indica que já existe uma tabela com o nome informado."""


class FormulaExcelError(ExcelAutomacaoError):
    """Representa falhas gerais relacionadas a fórmulas do Excel."""


class FormulaInvalidaError(FormulaExcelError, ValueError):
    """Indica que uma fórmula possui conteúdo ou sintaxe inválida."""


class FormatacaoExcelError(ExcelAutomacaoError):
    """Representa falhas durante a aplicação de estilos e formatos."""


# ----------------------------------------------------------------------------
# EXCEÇÕES DE IMPORTAÇÃO, EXPORTAÇÃO E OPERAÇÕES
# ----------------------------------------------------------------------------

class ImportacaoExcelError(ExcelAutomacaoError):
    """Representa falhas durante a importação de dados do Excel."""


class ExportacaoExcelError(ExcelAutomacaoError):
    """Representa falhas durante a exportação de dados para o Excel."""


class OperacaoExcelError(ExcelAutomacaoError):
    """Representa uma operação Excel que não pôde ser concluída."""


# ----------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ----------------------------------------------------------------------------

def criar_erro_excel(
    classe_erro: type[ExcelAutomacaoError],
    mensagem: str,
    contexto: dict[str, Any] | None = None,
    erro_original: Exception | None = None,
) -> ExcelAutomacaoError:
    """
    Cria uma exceção personalizada da hierarquia Excel.

    Args:
        classe_erro (type[ExcelAutomacaoError]):
            Classe de exceção que será instanciada.
        mensagem (str):
            Mensagem principal do erro.
        contexto (dict[str, Any] | None):
            Informações adicionais da operação.
        erro_original (Exception | None):
            Exceção original associada à falha.

    Returns:
        ExcelAutomacaoError:
            Instância da exceção solicitada.

    Raises:
        TypeError:
            Caso classe_erro não seja uma classe da hierarquia Excel.
    """

    # Verifica se o valor informado é uma classe
    if not isinstance(classe_erro, type):
        # Interrompe a execução quando o parâmetro não é uma classe
        raise TypeError(
            "O parâmetro 'classe_erro' deve ser uma classe de exceção."
        )

    # Verifica se a classe pertence à hierarquia de exceções Excel
    if not issubclass(classe_erro, ExcelAutomacaoError):
        # Interrompe a execução quando a classe não pertence ao módulo
        raise TypeError(
            "A classe informada deve herdar de ExcelAutomacaoError."
        )

    # Cria a exceção com os dados padronizados
    erro = classe_erro(
        mensagem,
        contexto=contexto,
        erro_original=erro_original,
    )

    # Retorna a exceção criada
    return erro


def converter_excecao(
    erro_original: Exception,
    classe_erro: type[ExcelAutomacaoError],
    mensagem: str,
    contexto: dict[str, Any] | None = None,
) -> ExcelAutomacaoError:
    """
    Converte uma exceção externa em uma exceção da biblioteca Excel.

    Args:
        erro_original (Exception):
            Exceção que originou a falha.
        classe_erro (type[ExcelAutomacaoError]):
            Classe personalizada que será criada.
        mensagem (str):
            Mensagem utilizada na nova exceção.
        contexto (dict[str, Any] | None):
            Informações adicionais relacionadas à operação.

    Returns:
        ExcelAutomacaoError:
            Exceção personalizada contendo a causa original.

    Raises:
        TypeError:
            Caso erro_original não seja uma exceção.
    """

    # Verifica se a causa recebida é uma exceção válida
    if not isinstance(erro_original, Exception):
        # Interrompe a execução quando a causa possui tipo inválido
        raise TypeError(
            "O parâmetro 'erro_original' deve ser uma exceção."
        )

    # Cria a nova exceção preservando a falha original
    erro_convertido = criar_erro_excel(
        classe_erro=classe_erro,
        mensagem=mensagem,
        contexto=contexto,
        erro_original=erro_original,
    )

    # Retorna a exceção convertida
    return erro_convertido


def obter_causa_raiz(erro: Exception) -> Exception:
    """
    Localiza a causa mais profunda de uma cadeia de exceções.

    A função considera erro_original, __cause__ e __context__, nesta ordem.
    Também impede ciclos quando uma exceção referencia a si mesma.

    Args:
        erro (Exception):
            Exceção cuja causa raiz será localizada.

    Returns:
        Exception:
            Exceção mais profunda encontrada na cadeia.

    Raises:
        TypeError:
            Caso o objeto informado não seja uma exceção.
    """

    # Verifica se o objeto recebido é uma exceção válida
    if not isinstance(erro, Exception):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'erro' deve ser uma exceção."
        )

    # Inicia a busca pela própria exceção recebida
    causa_atual = erro

    # Armazena os identificadores já visitados para impedir ciclos
    visitados: set[int] = set()

    # Percorre a cadeia enquanto existirem causas não visitadas
    while id(causa_atual) not in visitados:
        # Registra a exceção atual como visitada
        visitados.add(id(causa_atual))

        # Prioriza a causa personalizada armazenada pelo módulo
        erro_original = getattr(
            causa_atual,
            "erro_original",
            None,
        )

        # Seleciona a próxima causa disponível
        proxima_causa = (
            erro_original
            or causa_atual.__cause__
            or causa_atual.__context__
        )

        # Interrompe a busca quando não existe uma causa adicional
        if not isinstance(proxima_causa, Exception):
            break

        # Avança para a próxima exceção da cadeia
        causa_atual = proxima_causa

    # Retorna a causa mais profunda localizada
    return causa_atual


def excecao_para_dicionario(erro: Exception) -> dict[str, Any]:
    """
    Converte uma exceção comum ou personalizada em dicionário.

    Args:
        erro (Exception):
            Exceção que será convertida.

    Returns:
        dict[str, Any]:
            Estrutura adequada para logs e relatórios de execução.

    Raises:
        TypeError:
            Caso o objeto informado não seja uma exceção.
    """

    # Verifica se o objeto recebido é uma exceção
    if not isinstance(erro, Exception):
        # Interrompe a execução quando o tipo é inválido
        raise TypeError(
            "O parâmetro 'erro' deve ser uma exceção."
        )

    # Utiliza a conversão especializada quando disponível
    if isinstance(erro, ExcelAutomacaoError):
        return erro.para_dicionario()

    # Cria uma estrutura básica para exceções externas
    dados = {
        "tipo": erro.__class__.__name__,
        "mensagem": str(erro),
        "contexto": {},
        "erro_original": None,
    }

    # Retorna os dados padronizados
    return dados


# ----------------------------------------------------------------------------
# EXPORTAÇÕES PÚBLICAS
# ----------------------------------------------------------------------------

# Define explicitamente os objetos públicos disponibilizados pelo módulo
__all__ = [
    "ExcelAutomacaoError",
    "ExcelValidacaoError",
    "DadosExcelError",
    "DadosObrigatoriosError",
    "ArquivoExcelError",
    "ArquivoExcelInvalidoError",
    "ArquivoExcelNaoEncontradoError",
    "ArquivoExcelExistenteError",
    "WorkbookError",
    "WorkbookAberturaError",
    "WorkbookSalvamentoError",
    "WorkbookFechamentoError",
    "PlanilhaError",
    "PlanilhaInvalidaError",
    "PlanilhaNaoEncontradaError",
    "PlanilhaExistenteError",
    "NomePlanilhaInvalidoError",
    "CelulaError",
    "CelulaInvalidaError",
    "IntervaloError",
    "IntervaloInvalidoError",
    "ColunaError",
    "ColunaInvalidaError",
    "ColunaNaoEncontradaError",
    "LinhaError",
    "LinhaInvalidaError",
    "TabelaExcelError",
    "TabelaNaoEncontradaError",
    "TabelaExistenteError",
    "FormulaExcelError",
    "FormulaInvalidaError",
    "FormatacaoExcelError",
    "ImportacaoExcelError",
    "ExportacaoExcelError",
    "OperacaoExcelError",
    "criar_erro_excel",
    "converter_excecao",
    "obter_causa_raiz",
    "excecao_para_dicionario",
]
