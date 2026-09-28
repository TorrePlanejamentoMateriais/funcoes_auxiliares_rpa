"""
============================================================
Modulo: OUTLOOK / anexos.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Biblioteca de funcoes auxiliares para consultar, validar, adicionar,
    salvar e remover anexos de itens do Outlook classico via COM.
Funcoes disponiveis:
    - obter_colecao_anexos()
    - contar_anexos()
    - email_possui_anexos()
    - listar_anexos()
    - obter_nomes_anexos()
    - obter_anexo_por_indice()
    - obter_anexo_por_nome()
    - anexo_existe()
    - validar_extensao_anexo()
    - validar_tamanho_anexo()
    - validar_anexo()
    - preparar_diretorio_anexos()
    - sanitizar_nome_arquivo()
    - gerar_caminho_unico()
    - salvar_anexo()
    - salvar_anexo_por_indice()
    - salvar_anexo_por_nome()
    - salvar_todos_anexos()
    - salvar_anexos_por_extensao()
    - adicionar_anexo()
    - adicionar_anexos()
    - remover_anexo()
    - remover_todos_anexos()
Dependencias:
    - Outlook classico para Windows
    - pywin32 durante a execucao real da automacao
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de consulta e validacao de anexos.
        - Inclusao de salvamento com tratamento de conflitos.
        - Inclusao de adicao e remocao de anexos em mensagens.
============================================================
"""

# ----------------------------------------------------------------------------
# IMPORTACOES
# ----------------------------------------------------------------------------

# Importa re para remover caracteres invalidos dos nomes dos arquivos
import re

# Importa shutil para apoiar validacoes e operacoes futuras de arquivos
import shutil

# Importa Path para manipular caminhos de forma segura
from pathlib import Path

# Importa Any para tipar objetos COM dinamicos do Outlook
from typing import Any

# Importa a excecao principal da biblioteca quando ela esta disponivel
try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base utilizada quando o modulo central nao esta disponivel."""


# ----------------------------------------------------------------------------
# CONSTANTES
# ----------------------------------------------------------------------------

# Define caracteres proibidos em nomes de arquivos no Windows
PADRAO_CARACTERES_INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# Define nomes reservados pelo sistema operacional Windows
NOMES_RESERVADOS_WINDOWS = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{numero}" for numero in range(1, 10)),
    *(f"LPT{numero}" for numero in range(1, 10)),
}

# Define o tipo de anexo incorporado por valor no Outlook
OL_BY_VALUE = 1


# ----------------------------------------------------------------------------
# EXCECOES DO MODULO
# ----------------------------------------------------------------------------

class OutlookAnexosError(AutomacaoError):
    """Erro base para operacoes de anexos do Outlook."""


class ItemOutlookInvalidoError(OutlookAnexosError):
    """O item informado nao possui uma colecao de anexos valida."""


class AnexoNaoEncontradoError(OutlookAnexosError):
    """O anexo solicitado nao foi encontrado no item do Outlook."""


class AnexoInvalidoError(OutlookAnexosError):
    """O anexo nao atende aos criterios de validacao definidos."""


class AnexoNaoSalvoError(OutlookAnexosError):
    """O anexo nao pode ser salvo no sistema de arquivos."""


class AnexoNaoAdicionadoError(OutlookAnexosError):
    """O arquivo nao pode ser adicionado como anexo ao item."""


class AnexoNaoRemovidoError(OutlookAnexosError):
    """O anexo nao pode ser removido do item do Outlook."""


class ArquivoAnexoExistenteError(OutlookAnexosError):
    """Ja existe um arquivo no caminho definido para o anexo."""


# ----------------------------------------------------------------------------
# FUNCOES AUXILIARES INTERNAS
# ----------------------------------------------------------------------------

def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """
    Obtem uma propriedade de um objeto COM sem interromper a automacao.

    Args:
        objeto (Any):
            Objeto COM ou mock cuja propriedade sera consultada.
        nome (str):
            Nome da propriedade que sera obtida.
        padrao (Any):
            Valor retornado quando a propriedade nao estiver disponivel.

    Returns:
        Any:
            Valor da propriedade ou o valor padrao.
    """
    # Tenta acessar a propriedade dinamica do objeto COM
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_item_outlook(item: Any) -> None:
    """
    Valida se um item expoe a colecao Attachments do Outlook.

    Args:
        item (Any):
            MailItem, AppointmentItem ou outro item compativel.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        ItemOutlookInvalidoError:
            Caso o item seja None ou nao possua Attachments.
    """
    # Rejeita objetos ausentes antes de acessar suas propriedades
    if item is None:
        raise ItemOutlookInvalidoError(
            "O item do Outlook nao pode ser None."
        )

    # Verifica se o objeto expoe a colecao de anexos
    try:
        anexos = item.Attachments
    except Exception as error:
        raise ItemOutlookInvalidoError(
            "O item informado nao possui uma colecao Attachments valida."
        ) from error

    # Rejeita uma colecao ausente
    if anexos is None:
        raise ItemOutlookInvalidoError(
            "A colecao Attachments do item nao esta disponivel."
        )


def _validar_indice(indice: int, quantidade: int | None = None) -> None:
    """
    Valida um indice baseado em um utilizado pela colecao COM do Outlook.

    Args:
        indice (int):
            Posicao do anexo na colecao, iniciando em um.
        quantidade (int | None):
            Quantidade maxima disponivel na colecao.

    Returns:
        None:
            A funcao apenas realiza a validacao.

    Raises:
        TypeError:
            Caso indice nao seja inteiro.
        ValueError:
            Caso indice seja menor que um.
        AnexoNaoEncontradoError:
            Caso indice seja superior a quantidade disponivel.
    """
    # Rejeita booleanos porque bool tambem e uma subclasse de int
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError("O parametro 'indice' deve ser um numero inteiro.")

    # A colecao Attachments utiliza indices iniciados em um
    if indice < 1:
        raise ValueError("O parametro 'indice' deve ser maior ou igual a 1.")

    # Valida o limite quando a quantidade foi informada
    if quantidade is not None and indice > quantidade:
        raise AnexoNaoEncontradoError(
            f"Nao existe anexo no indice {indice}. Quantidade: {quantidade}."
        )


def _normalizar_extensoes(
    extensoes: str | list[str] | tuple[str, ...] | set[str] | None,
) -> tuple[str, ...] | None:
    """
    Normaliza uma extensao ou colecao de extensoes.

    Args:
        extensoes (str | list | tuple | set | None):
            Extensoes utilizadas nos filtros e validacoes.

    Returns:
        tuple[str, ...] | None:
            Extensoes em letras minusculas e iniciadas por ponto.

    Raises:
        TypeError:
            Caso o parametro possua formato invalido.
        ValueError:
            Caso a colecao esteja vazia.
    """
    # Mantem None quando nenhum filtro foi solicitado
    if extensoes is None:
        return None

    # Converte uma extensao unica para uma colecao
    if isinstance(extensoes, str):
        extensoes = [extensoes]

    # Valida os tipos de colecao aceitos
    if not isinstance(extensoes, (list, tuple, set)):
        raise TypeError(
            "O parametro 'extensoes' deve ser string, lista, tupla, set ou None."
        )

    # Rejeita colecoes vazias
    if not extensoes:
        raise ValueError("O parametro 'extensoes' nao pode estar vazio.")

    # Normaliza cada extensao
    resultado = []
    for extensao in extensoes:
        if not isinstance(extensao, str) or not extensao.strip():
            raise TypeError("Todas as extensoes devem ser strings nao vazias.")

        valor = extensao.strip().lower()
        if not valor.startswith("."):
            valor = f".{valor}"
        resultado.append(valor)

    # Remove duplicidades preservando a ordem
    return tuple(dict.fromkeys(resultado))


def _validar_tamanho(tamanho: int, nome: str) -> None:
    """
    Valida um tamanho de arquivo expresso em bytes.

    Args:
        tamanho (int):
            Valor em bytes que sera validado.
        nome (str):
            Nome utilizado na mensagem de erro.

    Returns:
        None:
            A funcao apenas realiza a validacao.
    """
    # Rejeita booleanos e valores nao inteiros
    if isinstance(tamanho, bool) or not isinstance(tamanho, int):
        raise TypeError(f"O parametro '{nome}' deve ser um numero inteiro.")

    # Tamanhos negativos nao representam arquivos validos
    if tamanho < 0:
        raise ValueError(f"O parametro '{nome}' nao pode ser negativo.")


# ----------------------------------------------------------------------------
# CONSULTA DE ANEXOS
# ----------------------------------------------------------------------------

def obter_colecao_anexos(item: Any) -> Any:
    """
    Retorna a colecao Attachments de um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que sera consultado.

    Returns:
        Any:
            Colecao COM Attachments do item.
    """
    # Valida o item antes de acessar a colecao
    _validar_item_outlook(item)

    # Retorna a colecao dinamica do Outlook
    return item.Attachments


def contar_anexos(item: Any) -> int:
    """
    Retorna a quantidade de anexos existentes no item.

    Args:
        item (Any):
            Item do Outlook que sera consultado.

    Returns:
        int:
            Quantidade atual de anexos.
    """
    # Recupera a colecao validada
    anexos = obter_colecao_anexos(item)

    try:
        # A propriedade Count informa a quantidade de itens COM
        return int(anexos.Count)
    except Exception as error:
        raise ItemOutlookInvalidoError(
            "Nao foi possivel contar os anexos do item."
        ) from error


def email_possui_anexos(item: Any) -> bool:
    """
    Verifica se o item do Outlook possui ao menos um anexo.

    Args:
        item (Any):
            Item do Outlook que sera consultado.

    Returns:
        bool:
            True quando existir anexo e False caso contrario.
    """
    # Compara a quantidade atual com zero
    return contar_anexos(item) > 0


def listar_anexos(item: Any) -> list[dict[str, Any]]:
    """
    Retorna os metadados dos anexos de um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que sera consultado.

    Returns:
        list[dict[str, Any]]:
            Indice, nome, nome exibido, tamanho, tipo e posicao dos anexos.

    Raises:
        OutlookAnexosError:
            Caso algum anexo nao possa ser consultado.
    """
    # Recupera a colecao e sua quantidade
    anexos = obter_colecao_anexos(item)
    quantidade = contar_anexos(item)
    resultado = []

    # A colecao COM do Outlook utiliza indices iniciados em um
    for indice in range(1, quantidade + 1):
        try:
            anexo = anexos.Item(indice)
            resultado.append(
                {
                    "indice": indice,
                    "nome": str(_obter_propriedade(anexo, "FileName", "")),
                    "nome_exibicao": str(
                        _obter_propriedade(anexo, "DisplayName", "")
                    ),
                    "tamanho": int(_obter_propriedade(anexo, "Size", 0) or 0),
                    "tipo": _obter_propriedade(anexo, "Type"),
                    "posicao": _obter_propriedade(anexo, "Position"),
                }
            )
        except Exception as error:
            raise OutlookAnexosError(
                f"Nao foi possivel consultar o anexo no indice {indice}."
            ) from error

    # Retorna os metadados coletados
    return resultado


def obter_nomes_anexos(item: Any) -> list[str]:
    """
    Retorna somente os nomes dos arquivos anexados.

    Args:
        item (Any):
            Item do Outlook que sera consultado.

    Returns:
        list[str]:
            Nomes dos anexos na ordem da colecao.
    """
    # Extrai a propriedade nome dos metadados
    return [dados["nome"] for dados in listar_anexos(item)]


def obter_anexo_por_indice(item: Any, indice: int) -> Any:
    """
    Retorna o anexo localizado em determinado indice COM.

    Args:
        item (Any):
            Item do Outlook que contem o anexo.
        indice (int):
            Posicao iniciada em um.

    Returns:
        Any:
            Objeto COM Attachment localizado.
    """
    # Valida o indice em relacao a quantidade disponivel
    quantidade = contar_anexos(item)
    _validar_indice(indice, quantidade)

    try:
        # Retorna o objeto Attachment da colecao
        return obter_colecao_anexos(item).Item(indice)
    except Exception as error:
        raise AnexoNaoEncontradoError(
            f"Nao foi possivel obter o anexo no indice {indice}."
        ) from error


def obter_anexo_por_nome(
    item: Any,
    nome: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
) -> Any:
    """
    Localiza um anexo por nome completo ou parcial.

    Args:
        item (Any):
            Item do Outlook que sera pesquisado.
        nome (str):
            Nome completo ou trecho procurado.
        correspondencia_exata (bool):
            Define se o nome deve corresponder exatamente.
        considerar_maiusculas (bool):
            Define se a pesquisa diferencia maiusculas de minusculas.

    Returns:
        Any:
            Primeiro Attachment correspondente.
    """
    # Valida o nome e as opcoes de pesquisa
    if not isinstance(nome, str) or not nome.strip():
        raise TypeError("O parametro 'nome' deve ser uma string nao vazia.")

    if not isinstance(correspondencia_exata, bool):
        raise TypeError("O parametro 'correspondencia_exata' deve ser booleano.")

    if not isinstance(considerar_maiusculas, bool):
        raise TypeError("O parametro 'considerar_maiusculas' deve ser booleano.")

    # Normaliza o valor procurado quando necessario
    procurado = nome.strip()
    if not considerar_maiusculas:
        procurado = procurado.lower()

    # Percorre os anexos na ordem original
    for indice in range(1, contar_anexos(item) + 1):
        anexo = obter_anexo_por_indice(item, indice)
        atual = str(_obter_propriedade(anexo, "FileName", ""))

        if not considerar_maiusculas:
            atual = atual.lower()

        corresponde = atual == procurado if correspondencia_exata else procurado in atual
        if corresponde:
            return anexo

    # Gera erro quando nenhuma correspondencia e localizada
    raise AnexoNaoEncontradoError(
        f"Nenhum anexo correspondente a '{nome}' foi encontrado."
    )


def anexo_existe(
    item: Any,
    nome: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
) -> bool:
    """
    Verifica se existe um anexo correspondente ao nome informado.

    Args:
        item (Any):
            Item do Outlook que sera pesquisado.
        nome (str):
            Nome completo ou parcial procurado.
        correspondencia_exata (bool):
            Define se a comparacao sera exata.
        considerar_maiusculas (bool):
            Define se a comparacao diferencia maiusculas.

    Returns:
        bool:
            True quando o anexo existir e False caso contrario.
    """
    # Tenta localizar o anexo e converte a ausencia em False
    try:
        obter_anexo_por_nome(
            item,
            nome,
            correspondencia_exata,
            considerar_maiusculas,
        )
        return True
    except AnexoNaoEncontradoError:
        return False


# ----------------------------------------------------------------------------
# VALIDACAO DE ANEXOS
# ----------------------------------------------------------------------------

def validar_extensao_anexo(
    anexo: Any,
    extensoes_permitidas: str | list[str] | tuple[str, ...] | set[str],
) -> bool:
    """
    Valida se a extensao do anexo pertence a lista permitida.

    Args:
        anexo (Any):
            Objeto Attachment que sera validado.
        extensoes_permitidas (str | list | tuple | set):
            Extensoes aceitas.

    Returns:
        bool:
            True quando a extensao for permitida.

    Raises:
        AnexoInvalidoError:
            Caso a extensao nao esteja permitida.
    """
    # Normaliza as extensoes e recupera o nome do arquivo
    extensoes = _normalizar_extensoes(extensoes_permitidas)
    nome = str(_obter_propriedade(anexo, "FileName", ""))
    extensao = Path(nome).suffix.lower()

    # Verifica se a extensao pertence ao conjunto aceito
    if extensao not in extensoes:
        raise AnexoInvalidoError(
            f"A extensao '{extensao}' do anexo '{nome}' nao e permitida."
        )

    return True


def validar_tamanho_anexo(
    anexo: Any,
    tamanho_minimo: int = 0,
    tamanho_maximo: int | None = None,
) -> bool:
    """
    Valida o tamanho do anexo em bytes.

    Args:
        anexo (Any):
            Objeto Attachment que sera validado.
        tamanho_minimo (int):
            Tamanho minimo permitido em bytes.
        tamanho_maximo (int | None):
            Tamanho maximo permitido em bytes.

    Returns:
        bool:
            True quando o tamanho estiver dentro dos limites.
    """
    # Valida os limites recebidos
    _validar_tamanho(tamanho_minimo, "tamanho_minimo")

    if tamanho_maximo is not None:
        _validar_tamanho(tamanho_maximo, "tamanho_maximo")
        if tamanho_maximo < tamanho_minimo:
            raise ValueError(
                "O tamanho maximo nao pode ser menor que o tamanho minimo."
            )

    # Recupera o tamanho disponibilizado pelo objeto COM
    tamanho = int(_obter_propriedade(anexo, "Size", 0) or 0)
    nome = str(_obter_propriedade(anexo, "FileName", ""))

    # Valida o limite inferior
    if tamanho < tamanho_minimo:
        raise AnexoInvalidoError(
            f"O anexo '{nome}' possui {tamanho} bytes, abaixo do minimo."
        )

    # Valida o limite superior quando ele foi definido
    if tamanho_maximo is not None and tamanho > tamanho_maximo:
        raise AnexoInvalidoError(
            f"O anexo '{nome}' possui {tamanho} bytes, acima do maximo."
        )

    return True


def validar_anexo(
    anexo: Any,
    extensoes_permitidas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    tamanho_minimo: int = 0,
    tamanho_maximo: int | None = None,
) -> bool:
    """
    Valida extensao e tamanho de um anexo do Outlook.

    Args:
        anexo (Any):
            Objeto Attachment que sera validado.
        extensoes_permitidas (str | list | tuple | set | None):
            Extensoes aceitas ou None para nao filtrar.
        tamanho_minimo (int):
            Tamanho minimo permitido em bytes.
        tamanho_maximo (int | None):
            Tamanho maximo permitido em bytes.

    Returns:
        bool:
            True quando todas as validacoes forem atendidas.
    """
    # Valida a extensao somente quando um filtro foi informado
    if extensoes_permitidas is not None:
        validar_extensao_anexo(anexo, extensoes_permitidas)

    # Valida os limites de tamanho
    validar_tamanho_anexo(anexo, tamanho_minimo, tamanho_maximo)
    return True


# ----------------------------------------------------------------------------
# CAMINHOS E NOMES
# ----------------------------------------------------------------------------

def preparar_diretorio_anexos(diretorio: str | Path) -> Path:
    """
    Cria e retorna o diretorio utilizado para salvar anexos.

    Args:
        diretorio (str | Path):
            Diretorio de destino dos arquivos.

    Returns:
        Path:
            Caminho absoluto do diretorio preparado.
    """
    # Valida o tipo do caminho recebido
    if not isinstance(diretorio, (str, Path)):
        raise TypeError("O parametro 'diretorio' deve ser uma string ou Path.")

    # Converte o caminho para absoluto
    caminho = Path(diretorio).expanduser().resolve()

    # Rejeita um arquivo existente usado como diretorio
    if caminho.exists() and not caminho.is_dir():
        raise NotADirectoryError(f"O caminho nao e um diretorio: {caminho}")

    # Cria o diretorio e seus pais quando necessario
    caminho.mkdir(parents=True, exist_ok=True)
    return caminho


def sanitizar_nome_arquivo(nome: str, tamanho_maximo: int = 180) -> str:
    """
    Remove caracteres invalidos de um nome de arquivo de anexo.

    Args:
        nome (str):
            Nome original do anexo.
        tamanho_maximo (int):
            Quantidade maxima de caracteres do nome final.

    Returns:
        str:
            Nome seguro para gravacao no Windows.
    """
    # Valida o nome e o limite
    if not isinstance(nome, str) or not nome.strip():
        raise TypeError("O parametro 'nome' deve ser uma string nao vazia.")

    if isinstance(tamanho_maximo, bool) or not isinstance(tamanho_maximo, int):
        raise TypeError("O parametro 'tamanho_maximo' deve ser inteiro.")

    if tamanho_maximo <= 0:
        raise ValueError("O tamanho maximo deve ser maior que zero.")

    # Separa o sufixo para preservar a extensao durante o corte
    caminho_nome = Path(nome.strip())
    extensao = caminho_nome.suffix
    base = caminho_nome.stem if extensao else caminho_nome.name

    # Substitui caracteres proibidos e remove pontos ou espacos finais
    base = PADRAO_CARACTERES_INVALIDOS.sub("_", base).strip(" ._")
    extensao = PADRAO_CARACTERES_INVALIDOS.sub("_", extensao).strip(" ")

    # Usa um nome padrao quando nenhum caractere valido permanece
    if not base:
        base = "anexo"

    # Evita nomes reservados pelo Windows
    if base.upper() in NOMES_RESERVADOS_WINDOWS:
        base = f"_{base}"

    # Preserva espaco suficiente para a extensao
    limite_base = max(1, tamanho_maximo - len(extensao))
    return f"{base[:limite_base]}{extensao}"


def gerar_caminho_unico(
    diretorio: str | Path,
    nome_arquivo: str,
    separador: str = "_",
) -> Path:
    """
    Gera um caminho inexistente acrescentando um contador ao nome.

    Args:
        diretorio (str | Path):
            Diretorio onde o arquivo sera salvo.
        nome_arquivo (str):
            Nome desejado para o arquivo.
        separador (str):
            Texto inserido antes do contador.

    Returns:
        Path:
            Primeiro caminho disponivel encontrado.
    """
    # Prepara o diretorio e sanitiza o nome
    destino = preparar_diretorio_anexos(diretorio)
    nome = sanitizar_nome_arquivo(nome_arquivo)

    if not isinstance(separador, str):
        raise TypeError("O parametro 'separador' deve ser uma string.")

    # Retorna o caminho original quando ele ainda nao existe
    candidato = destino / nome
    if not candidato.exists():
        return candidato

    # Acrescenta um contador ate encontrar um caminho livre
    base = Path(nome).stem
    extensao = Path(nome).suffix
    contador = 1

    while True:
        candidato = destino / f"{base}{separador}{contador}{extensao}"
        if not candidato.exists():
            return candidato
        contador += 1


# ----------------------------------------------------------------------------
# SALVAMENTO DE ANEXOS
# ----------------------------------------------------------------------------

def salvar_anexo(
    anexo: Any,
    diretorio: str | Path,
    nome_arquivo: str | None = None,
    sobrescrever: bool = False,
    gerar_nome_unico: bool = True,
    extensoes_permitidas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    tamanho_minimo: int = 0,
    tamanho_maximo: int | None = None,
) -> Path:
    """
    Valida e salva um Attachment do Outlook no sistema de arquivos.

    Args:
        anexo (Any):
            Objeto COM Attachment que sera salvo.
        diretorio (str | Path):
            Diretorio de destino.
        nome_arquivo (str | None):
            Nome opcional aplicado ao arquivo salvo.
        sobrescrever (bool):
            Define se um arquivo existente podera ser substituido.
        gerar_nome_unico (bool):
            Acrescenta contador quando houver conflito de nome.
        extensoes_permitidas (str | list | tuple | set | None):
            Extensoes aceitas para salvamento.
        tamanho_minimo (int):
            Tamanho minimo permitido em bytes.
        tamanho_maximo (int | None):
            Tamanho maximo permitido em bytes.

    Returns:
        Path:
            Caminho absoluto do arquivo salvo.
    """
    # Valida as opcoes booleanas
    if not isinstance(sobrescrever, bool) or not isinstance(gerar_nome_unico, bool):
        raise TypeError(
            "Os parametros 'sobrescrever' e 'gerar_nome_unico' devem ser booleanos."
        )

    # Executa as validacoes antes de gravar o arquivo
    validar_anexo(
        anexo,
        extensoes_permitidas,
        tamanho_minimo,
        tamanho_maximo,
    )

    # Define o nome recebido ou o nome original do anexo
    nome_original = str(_obter_propriedade(anexo, "FileName", "anexo"))
    nome_final = nome_original if nome_arquivo is None else nome_arquivo
    nome_final = sanitizar_nome_arquivo(nome_final)
    destino = preparar_diretorio_anexos(diretorio) / nome_final

    # Resolve conflitos conforme as opcoes informadas
    if destino.exists():
        if sobrescrever:
            if destino.is_dir():
                raise ArquivoAnexoExistenteError(
                    f"O destino representa um diretorio: {destino}"
                )
            destino.unlink()
        elif gerar_nome_unico:
            destino = gerar_caminho_unico(destino.parent, destino.name)
        else:
            raise ArquivoAnexoExistenteError(
                f"Ja existe um arquivo no destino: {destino}"
            )

    try:
        # SaveAsFile recebe obrigatoriamente o caminho completo do destino
        anexo.SaveAsFile(str(destino))
    except Exception as error:
        raise AnexoNaoSalvoError(
            f"Nao foi possivel salvar o anexo em: {destino}"
        ) from error

    # Confirma que o Outlook efetivamente criou o arquivo
    if not destino.is_file():
        raise AnexoNaoSalvoError(
            f"O Outlook nao criou o arquivo esperado: {destino}"
        )

    return destino


def salvar_anexo_por_indice(
    item: Any,
    indice: int,
    diretorio: str | Path,
    **opcoes: Any,
) -> Path:
    """
    Localiza um anexo pelo indice e salva o arquivo.

    Args:
        item (Any):
            Item do Outlook que contem o anexo.
        indice (int):
            Posicao do anexo iniciada em um.
        diretorio (str | Path):
            Diretorio de destino.
        **opcoes (Any):
            Opcoes encaminhadas para salvar_anexo.

    Returns:
        Path:
            Caminho do arquivo salvo.
    """
    # Localiza o objeto Attachment e delega o salvamento
    anexo = obter_anexo_por_indice(item, indice)
    return salvar_anexo(anexo, diretorio, **opcoes)


def salvar_anexo_por_nome(
    item: Any,
    nome: str,
    diretorio: str | Path,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    **opcoes: Any,
) -> Path:
    """
    Localiza um anexo pelo nome e salva o arquivo.

    Args:
        item (Any):
            Item do Outlook que contem o anexo.
        nome (str):
            Nome completo ou trecho procurado.
        diretorio (str | Path):
            Diretorio de destino.
        correspondencia_exata (bool):
            Define se o nome deve corresponder exatamente.
        considerar_maiusculas (bool):
            Define se a pesquisa diferencia maiusculas.
        **opcoes (Any):
            Opcoes encaminhadas para salvar_anexo.

    Returns:
        Path:
            Caminho do arquivo salvo.
    """
    # Localiza o objeto Attachment correspondente
    anexo = obter_anexo_por_nome(
        item,
        nome,
        correspondencia_exata,
        considerar_maiusculas,
    )

    # Delega o salvamento para a funcao principal
    return salvar_anexo(anexo, diretorio, **opcoes)


def salvar_todos_anexos(
    item: Any,
    diretorio: str | Path,
    extensoes_permitidas: str | list[str] | tuple[str, ...] | set[str] | None = None,
    tamanho_minimo: int = 0,
    tamanho_maximo: int | None = None,
    ignorar_invalidos: bool = False,
    sobrescrever: bool = False,
    gerar_nome_unico: bool = True,
) -> list[Path]:
    """
    Salva todos os anexos validos existentes em um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que sera processado.
        diretorio (str | Path):
            Diretorio de destino.
        extensoes_permitidas (str | list | tuple | set | None):
            Extensoes aceitas ou None para aceitar todas.
        tamanho_minimo (int):
            Tamanho minimo permitido em bytes.
        tamanho_maximo (int | None):
            Tamanho maximo permitido em bytes.
        ignorar_invalidos (bool):
            Ignora anexos que nao atendem aos criterios.
        sobrescrever (bool):
            Define se arquivos existentes poderao ser substituidos.
        gerar_nome_unico (bool):
            Gera nomes alternativos quando houver conflito.

    Returns:
        list[Path]:
            Caminhos dos arquivos salvos.
    """
    # Valida a opcao de ignorar anexos invalidos
    if not isinstance(ignorar_invalidos, bool):
        raise TypeError("O parametro 'ignorar_invalidos' deve ser booleano.")

    # Inicializa a lista dos arquivos gerados
    arquivos_salvos = []

    # Percorre a colecao COM iniciada em um
    for indice in range(1, contar_anexos(item) + 1):
        anexo = obter_anexo_por_indice(item, indice)

        try:
            caminho = salvar_anexo(
                anexo=anexo,
                diretorio=diretorio,
                sobrescrever=sobrescrever,
                gerar_nome_unico=gerar_nome_unico,
                extensoes_permitidas=extensoes_permitidas,
                tamanho_minimo=tamanho_minimo,
                tamanho_maximo=tamanho_maximo,
            )
            arquivos_salvos.append(caminho)
        except AnexoInvalidoError:
            if not ignorar_invalidos:
                raise

    return arquivos_salvos


def salvar_anexos_por_extensao(
    item: Any,
    diretorio: str | Path,
    extensoes: str | list[str] | tuple[str, ...] | set[str],
    **opcoes: Any,
) -> list[Path]:
    """
    Salva somente anexos que possuem as extensoes informadas.

    Args:
        item (Any):
            Item do Outlook que sera processado.
        diretorio (str | Path):
            Diretorio de destino.
        extensoes (str | list | tuple | set):
            Extensoes aceitas no salvamento.
        **opcoes (Any):
            Opcoes adicionais encaminhadas para salvar_todos_anexos.

    Returns:
        list[Path]:
            Caminhos dos anexos correspondentes salvos.
    """
    # Delega o filtro de extensao para o salvamento em lote
    return salvar_todos_anexos(
        item=item,
        diretorio=diretorio,
        extensoes_permitidas=extensoes,
        ignorar_invalidos=True,
        **opcoes,
    )


# ----------------------------------------------------------------------------
# ADICAO E REMOCAO DE ANEXOS
# ----------------------------------------------------------------------------

def adicionar_anexo(
    item: Any,
    caminho_arquivo: str | Path,
    nome_exibicao: str | None = None,
    tipo: int = OL_BY_VALUE,
    posicao: int = 1,
) -> Any:
    """
    Adiciona um arquivo como anexo a um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que recebera o anexo.
        caminho_arquivo (str | Path):
            Arquivo existente que sera anexado.
        nome_exibicao (str | None):
            Nome opcional exibido no item do Outlook.
        tipo (int):
            Tipo de anexo da enumeracao OlAttachmentType.
        posicao (int):
            Posicao do anexo no corpo do item.

    Returns:
        Any:
            Objeto COM Attachment criado.
    """
    # Valida o item e o caminho do arquivo
    anexos = obter_colecao_anexos(item)

    if not isinstance(caminho_arquivo, (str, Path)):
        raise TypeError("O parametro 'caminho_arquivo' deve ser string ou Path.")

    arquivo = Path(caminho_arquivo).expanduser().resolve()
    if not arquivo.is_file():
        raise FileNotFoundError(f"O arquivo para anexo nao existe: {arquivo}")

    # Valida os parametros numericos utilizados pelo metodo Add
    if isinstance(tipo, bool) or not isinstance(tipo, int):
        raise TypeError("O parametro 'tipo' deve ser inteiro.")

    if isinstance(posicao, bool) or not isinstance(posicao, int):
        raise TypeError("O parametro 'posicao' deve ser inteiro.")

    if posicao < 0:
        raise ValueError("O parametro 'posicao' nao pode ser negativo.")

    # Valida o nome de exibicao quando ele e fornecido
    if nome_exibicao is not None and (
        not isinstance(nome_exibicao, str) or not nome_exibicao.strip()
    ):
        raise TypeError(
            "O parametro 'nome_exibicao' deve ser string nao vazia ou None."
        )

    try:
        # O metodo Add recebe fonte, tipo, posicao e nome de exibicao
        return anexos.Add(
            str(arquivo),
            tipo,
            posicao,
            nome_exibicao.strip() if nome_exibicao else arquivo.name,
        )
    except Exception as error:
        raise AnexoNaoAdicionadoError(
            f"Nao foi possivel adicionar o arquivo como anexo: {arquivo}"
        ) from error


def adicionar_anexos(
    item: Any,
    caminhos: list[str | Path] | tuple[str | Path, ...] | set[str | Path],
    ignorar_inexistentes: bool = False,
) -> list[Any]:
    """
    Adiciona varios arquivos como anexos a um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que recebera os anexos.
        caminhos (list | tuple | set):
            Caminhos dos arquivos que serao adicionados.
        ignorar_inexistentes (bool):
            Ignora caminhos que nao representam arquivos existentes.

    Returns:
        list[Any]:
            Objetos Attachment criados.
    """
    # Valida a colecao e a opcao de tolerancia
    if not isinstance(caminhos, (list, tuple, set)) or not caminhos:
        raise TypeError("O parametro 'caminhos' deve ser uma colecao nao vazia.")

    if not isinstance(ignorar_inexistentes, bool):
        raise TypeError("O parametro 'ignorar_inexistentes' deve ser booleano.")

    # Adiciona cada arquivo valido
    resultado = []
    for caminho in caminhos:
        arquivo = Path(caminho).expanduser().resolve()

        if not arquivo.is_file() and ignorar_inexistentes:
            continue

        resultado.append(adicionar_anexo(item, arquivo))

    return resultado


def remover_anexo(item: Any, indice: int) -> bool:
    """
    Remove um anexo do item utilizando seu indice COM.

    Args:
        item (Any):
            Item do Outlook que contem o anexo.
        indice (int):
            Posicao iniciada em um.

    Returns:
        bool:
            True quando a remocao for concluida.
    """
    # Localiza o anexo antes de executar a exclusao
    anexo = obter_anexo_por_indice(item, indice)

    try:
        # Delete remove o Attachment da colecao do item
        anexo.Delete()
        return True
    except Exception as error:
        raise AnexoNaoRemovidoError(
            f"Nao foi possivel remover o anexo no indice {indice}."
        ) from error


def remover_todos_anexos(item: Any) -> int:
    """
    Remove todos os anexos de um item do Outlook.

    Args:
        item (Any):
            Item do Outlook que sera alterado.

    Returns:
        int:
            Quantidade de anexos removidos.
    """
    # Registra a quantidade inicial para o retorno
    quantidade = contar_anexos(item)

    # Remove sempre o ultimo indice para evitar deslocamento da colecao
    for indice in range(quantidade, 0, -1):
        remover_anexo(item, indice)

    return quantidade
