"""
============================================================
Modulo: OUTLOOK / pastas.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes reutilizaveis para obter, navegar, criar, listar, copiar,
    mover e excluir pastas do Outlook classico via automacao COM.
Dependencias:
    - Outlook classico para Windows
    - constantes.py
    - excecoes.py
============================================================
"""

from __future__ import annotations

from typing import Any

try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base usada quando o modulo central nao esta disponivel."""

try:
    from constantes import OL_FOLDER_INBOX
except ImportError:
    OL_FOLDER_INBOX = 6


class OutlookPastasError(AutomacaoError):
    """Erro base para operacoes relacionadas a pastas do Outlook."""


class PastaInvalidaError(OutlookPastasError):
    """A pasta informada esta ausente ou possui propriedades invalidas."""


class PastaNaoEncontradaError(OutlookPastasError):
    """A pasta solicitada nao foi encontrada na hierarquia do Outlook."""


class PastaCriacaoError(OutlookPastasError):
    """A pasta nao pode ser criada no local solicitado."""


class PastaAcaoError(OutlookPastasError):
    """Uma acao de mover, copiar, excluir ou exibir a pasta falhou."""


class CaminhoPastaError(OutlookPastasError):
    """O caminho de pasta informado possui formato invalido."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM e retorna um valor padrao em erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise PastaInvalidaError(f"O parametro '{nome}' nao pode ser None.")


def _validar_texto(valor: str, nome: str, permitir_vazio: bool = False) -> str:
    """Valida e normaliza uma propriedade textual."""
    if not isinstance(permitir_vazio, bool):
        raise TypeError("O parametro 'permitir_vazio' deve ser booleano.")
    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")
    texto = valor.strip()
    if not texto and not permitir_vazio:
        raise ValueError(f"O parametro '{nome}' nao pode estar vazio.")
    return texto


def _validar_booleano(valor: bool, nome: str) -> None:
    """Valida uma opcao booleana."""
    if not isinstance(valor, bool):
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")


def _normalizar_caminho(caminho: str) -> list[str]:
    """Normaliza um caminho Outlook e retorna seus segmentos."""
    texto = _validar_texto(caminho, "caminho")
    texto = texto.replace("/", "\\")
    while texto.startswith("\\"):
        texto = texto[1:]
    partes = [parte.strip() for parte in texto.split("\\") if parte.strip()]
    if not partes:
        raise CaminhoPastaError("O caminho nao possui segmentos validos.")
    return partes


def _comparar_nome(
    nome_atual: str,
    nome_procurado: str,
    considerar_maiusculas: bool,
) -> bool:
    """Compara nomes de pastas com sensibilidade opcional a maiusculas."""
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")
    if considerar_maiusculas:
        return nome_atual == nome_procurado
    return nome_atual.lower() == nome_procurado.lower()


def obter_pasta_padrao(namespace: Any, tipo_pasta: int = OL_FOLDER_INBOX) -> Any:
    """Retorna uma pasta padrao utilizando NameSpace.GetDefaultFolder."""
    _validar_objeto(namespace, "namespace")
    if isinstance(tipo_pasta, bool) or not isinstance(tipo_pasta, int):
        raise TypeError("O parametro 'tipo_pasta' deve ser inteiro.")
    try:
        pasta = namespace.GetDefaultFolder(tipo_pasta)
    except Exception as error:
        raise PastaNaoEncontradaError(
            f"Nao foi possivel obter a pasta padrao de tipo {tipo_pasta}."
        ) from error
    if pasta is None:
        raise PastaNaoEncontradaError(
            f"A pasta padrao de tipo {tipo_pasta} nao foi encontrada."
        )
    return pasta


def obter_pastas_raiz(namespace: Any) -> Any:
    """Retorna a colecao de caixas postais e stores na raiz do Namespace."""
    _validar_objeto(namespace, "namespace")
    try:
        pastas = namespace.Folders
    except Exception as error:
        raise PastaInvalidaError(
            "Nao foi possivel obter as pastas raiz do Namespace."
        ) from error
    if pastas is None:
        raise PastaInvalidaError("A colecao de pastas raiz nao esta disponivel.")
    return pastas


def contar_subpastas(pasta: Any) -> int:
    """Retorna a quantidade de subpastas de uma pasta."""
    _validar_objeto(pasta, "pasta")
    try:
        return int(pasta.Folders.Count)
    except Exception as error:
        raise PastaInvalidaError("Nao foi possivel contar as subpastas.") from error


def obter_subpasta_por_indice(pasta: Any, indice: int) -> Any:
    """Retorna uma subpasta pelo indice COM iniciado em um."""
    _validar_objeto(pasta, "pasta")
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError("O parametro 'indice' deve ser inteiro.")
    if indice < 1:
        raise ValueError("O parametro 'indice' deve ser maior ou igual a 1.")
    quantidade = contar_subpastas(pasta)
    if indice > quantidade:
        raise PastaNaoEncontradaError(
            f"Nao existe subpasta no indice {indice}. Quantidade: {quantidade}."
        )
    try:
        return pasta.Folders.Item(indice)
    except Exception as error:
        raise PastaNaoEncontradaError(
            f"Nao foi possivel obter a subpasta no indice {indice}."
        ) from error


def obter_subpasta(
    pasta: Any,
    nome: str,
    considerar_maiusculas: bool = False,
) -> Any:
    """Localiza uma subpasta pelo nome."""
    _validar_objeto(pasta, "pasta")
    nome_validado = _validar_texto(nome, "nome")
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")

    for indice in range(1, contar_subpastas(pasta) + 1):
        subpasta = obter_subpasta_por_indice(pasta, indice)
        nome_atual = str(_obter_propriedade(subpasta, "Name", "") or "")
        if _comparar_nome(nome_atual, nome_validado, considerar_maiusculas):
            return subpasta

    raise PastaNaoEncontradaError(
        f"A subpasta '{nome_validado}' nao foi encontrada."
    )


def subpasta_existe(
    pasta: Any,
    nome: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """Verifica se uma subpasta existe pelo nome."""
    try:
        obter_subpasta(pasta, nome, considerar_maiusculas)
        return True
    except PastaNaoEncontradaError:
        return False


def listar_subpastas(
    pasta: Any,
    recursivo: bool = False,
    incluir_objeto: bool = True,
    nivel_inicial: int = 0,
) -> list[dict[str, Any]]:
    """Lista subpastas e metadados, opcionalmente de forma recursiva."""
    _validar_objeto(pasta, "pasta")
    _validar_booleano(recursivo, "recursivo")
    _validar_booleano(incluir_objeto, "incluir_objeto")
    if isinstance(nivel_inicial, bool) or not isinstance(nivel_inicial, int):
        raise TypeError("O parametro 'nivel_inicial' deve ser inteiro.")
    if nivel_inicial < 0:
        raise ValueError("O parametro 'nivel_inicial' nao pode ser negativo.")

    resultado: list[dict[str, Any]] = []
    for indice in range(1, contar_subpastas(pasta) + 1):
        subpasta = obter_subpasta_por_indice(pasta, indice)
        dados = obter_dados_pasta(subpasta)
        dados["indice"] = indice
        dados["nivel"] = nivel_inicial
        if incluir_objeto:
            dados["pasta"] = subpasta
        resultado.append(dados)
        if recursivo:
            resultado.extend(
                listar_subpastas(
                    subpasta,
                    recursivo=True,
                    incluir_objeto=incluir_objeto,
                    nivel_inicial=nivel_inicial + 1,
                )
            )
    return resultado


def obter_pasta_por_caminho(
    namespace: Any,
    caminho: str,
    considerar_maiusculas: bool = False,
) -> Any:
    """Navega por um caminho como Caixa Postal\\Entrada\\Processados."""
    _validar_objeto(namespace, "namespace")
    partes = _normalizar_caminho(caminho)
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")

    raiz_encontrada = None
    pastas_raiz = obter_pastas_raiz(namespace)
    try:
        quantidade = int(pastas_raiz.Count)
    except Exception as error:
        raise PastaInvalidaError("Nao foi possivel contar as pastas raiz.") from error

    for indice in range(1, quantidade + 1):
        try:
            raiz = pastas_raiz.Item(indice)
        except Exception as error:
            raise PastaInvalidaError(
                f"Nao foi possivel consultar a pasta raiz {indice}."
            ) from error
        nome = str(_obter_propriedade(raiz, "Name", "") or "")
        if _comparar_nome(nome, partes[0], considerar_maiusculas):
            raiz_encontrada = raiz
            break

    if raiz_encontrada is None:
        raise PastaNaoEncontradaError(
            f"A pasta raiz '{partes[0]}' nao foi encontrada."
        )

    pasta_atual = raiz_encontrada
    for nome_subpasta in partes[1:]:
        pasta_atual = obter_subpasta(
            pasta_atual,
            nome_subpasta,
            considerar_maiusculas=considerar_maiusculas,
        )
    return pasta_atual


def criar_subpasta(
    pasta_pai: Any,
    nome: str,
    tipo_pasta: int | None = None,
    retornar_existente: bool = False,
) -> Any:
    """Cria uma subpasta e opcionalmente retorna uma pasta existente."""
    _validar_objeto(pasta_pai, "pasta_pai")
    nome_validado = _validar_texto(nome, "nome")
    _validar_booleano(retornar_existente, "retornar_existente")
    if tipo_pasta is not None and (
        isinstance(tipo_pasta, bool) or not isinstance(tipo_pasta, int)
    ):
        raise TypeError("O parametro 'tipo_pasta' deve ser inteiro ou None.")

    if retornar_existente and subpasta_existe(pasta_pai, nome_validado):
        return obter_subpasta(pasta_pai, nome_validado)

    try:
        if tipo_pasta is None:
            return pasta_pai.Folders.Add(nome_validado)
        return pasta_pai.Folders.Add(nome_validado, tipo_pasta)
    except Exception as error:
        raise PastaCriacaoError(
            f"Nao foi possivel criar a subpasta '{nome_validado}'."
        ) from error


def criar_caminho_pastas(
    pasta_inicial: Any,
    caminho: str,
    retornar_existentes: bool = True,
) -> Any:
    """Cria uma hierarquia de subpastas e retorna a ultima pasta."""
    _validar_objeto(pasta_inicial, "pasta_inicial")
    _validar_booleano(retornar_existentes, "retornar_existentes")
    partes = _normalizar_caminho(caminho)
    pasta_atual = pasta_inicial
    for nome in partes:
        pasta_atual = criar_subpasta(
            pasta_atual,
            nome,
            retornar_existente=retornar_existentes,
        )
    return pasta_atual


def renomear_pasta(pasta: Any, novo_nome: str) -> Any:
    """Altera o nome de uma pasta e retorna o proprio objeto."""
    _validar_objeto(pasta, "pasta")
    nome = _validar_texto(novo_nome, "novo_nome")
    try:
        pasta.Name = nome
        return pasta
    except Exception as error:
        raise PastaAcaoError("Nao foi possivel renomear a pasta.") from error


def mover_pasta(pasta: Any, pasta_destino: Any) -> Any:
    """Move uma pasta para outra pasta e retorna a pasta movida."""
    _validar_objeto(pasta, "pasta")
    _validar_objeto(pasta_destino, "pasta_destino")
    try:
        return pasta.MoveTo(pasta_destino)
    except Exception as error:
        raise PastaAcaoError("Nao foi possivel mover a pasta.") from error


def copiar_pasta(pasta: Any, pasta_destino: Any) -> Any:
    """Copia uma pasta para outra pasta e retorna a copia."""
    _validar_objeto(pasta, "pasta")
    _validar_objeto(pasta_destino, "pasta_destino")
    try:
        return pasta.CopyTo(pasta_destino)
    except Exception as error:
        raise PastaAcaoError("Nao foi possivel copiar a pasta.") from error


def excluir_pasta(pasta: Any) -> bool:
    """Exclui uma pasta do Outlook."""
    _validar_objeto(pasta, "pasta")
    try:
        pasta.Delete()
        return True
    except Exception as error:
        raise PastaAcaoError("Nao foi possivel excluir a pasta.") from error


def exibir_pasta(pasta: Any) -> Any:
    """Exibe a pasta no Outlook e retorna o proprio objeto."""
    _validar_objeto(pasta, "pasta")
    try:
        pasta.Display()
        return pasta
    except Exception as error:
        raise PastaAcaoError("Nao foi possivel exibir a pasta.") from error


def contar_itens(pasta: Any) -> int:
    """Retorna a quantidade de itens contidos na pasta."""
    _validar_objeto(pasta, "pasta")
    try:
        return int(pasta.Items.Count)
    except Exception as error:
        raise PastaInvalidaError("Nao foi possivel contar os itens da pasta.") from error


def contar_itens_nao_lidos(pasta: Any) -> int:
    """Retorna a quantidade de itens nao lidos informada pela pasta."""
    _validar_objeto(pasta, "pasta")
    try:
        return int(pasta.UnReadItemCount)
    except Exception as error:
        raise PastaInvalidaError(
            "Nao foi possivel contar os itens nao lidos."
        ) from error


def listar_itens(
    pasta: Any,
    ordenar_por: str | None = None,
    ordem_decrescente: bool = False,
    limite: int | None = None,
) -> list[Any]:
    """Lista itens de uma pasta com ordenacao e limite opcionais."""
    _validar_objeto(pasta, "pasta")
    _validar_booleano(ordem_decrescente, "ordem_decrescente")
    if ordenar_por is not None:
        ordenar_por = _validar_texto(ordenar_por, "ordenar_por")
    if limite is not None:
        if isinstance(limite, bool) or not isinstance(limite, int):
            raise TypeError("O parametro 'limite' deve ser inteiro ou None.")
        if limite < 1:
            raise ValueError("O parametro 'limite' deve ser maior que zero.")
    try:
        itens = pasta.Items
        if ordenar_por is not None:
            itens.Sort(ordenar_por, ordem_decrescente)
        quantidade = int(itens.Count)
    except Exception as error:
        raise PastaInvalidaError("Nao foi possivel consultar os itens da pasta.") from error
    quantidade = quantidade if limite is None else min(quantidade, limite)
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            resultado.append(itens.Item(indice))
        except Exception as error:
            raise PastaInvalidaError(
                f"Nao foi possivel obter o item no indice {indice}."
            ) from error
    return resultado


def obter_dados_pasta(pasta: Any) -> dict[str, Any]:
    """Converte as principais propriedades da pasta para dicionario."""
    _validar_objeto(pasta, "pasta")
    store = _obter_propriedade(pasta, "Store")
    return {
        "entry_id": _obter_propriedade(pasta, "EntryID"),
        "nome": str(_obter_propriedade(pasta, "Name", "") or ""),
        "caminho": str(_obter_propriedade(pasta, "FolderPath", "") or ""),
        "descricao": str(_obter_propriedade(pasta, "Description", "") or ""),
        "tipo_item_padrao": _obter_propriedade(pasta, "DefaultItemType"),
        "classe_mensagem_padrao": str(
            _obter_propriedade(pasta, "DefaultMessageClass", "") or ""
        ),
        "quantidade_itens": contar_itens(pasta),
        "quantidade_nao_lidos": contar_itens_nao_lidos(pasta),
        "quantidade_subpastas": contar_subpastas(pasta),
        "store_id": _obter_propriedade(pasta, "StoreID"),
        "store_nome": str(_obter_propriedade(store, "DisplayName", "") or ""),
    }


def obter_pasta_pai(pasta: Any) -> Any:
    """Retorna a pasta pai ou o objeto Namespace/Store correspondente."""
    _validar_objeto(pasta, "pasta")
    pai = _obter_propriedade(pasta, "Parent")
    if pai is None:
        raise PastaNaoEncontradaError("A pasta nao possui um objeto pai acessivel.")
    return pai


def obter_caminho_pasta(pasta: Any) -> str:
    """Retorna o caminho completo da pasta."""
    _validar_objeto(pasta, "pasta")
    caminho = str(_obter_propriedade(pasta, "FolderPath", "") or "").strip()
    if not caminho:
        raise PastaInvalidaError("A pasta nao possui FolderPath disponivel.")
    return caminho


def localizar_pastas_por_nome(
    pasta_inicial: Any,
    nome: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    incluir_pasta_inicial: bool = False,
) -> list[Any]:
    """Pesquisa pastas recursivamente pelo nome completo ou parcial."""
    _validar_objeto(pasta_inicial, "pasta_inicial")
    procurado = _validar_texto(nome, "nome")
    _validar_booleano(correspondencia_exata, "correspondencia_exata")
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")
    _validar_booleano(incluir_pasta_inicial, "incluir_pasta_inicial")

    resultado: list[Any] = []

    def corresponde(pasta: Any) -> bool:
        atual = str(_obter_propriedade(pasta, "Name", "") or "")
        if not considerar_maiusculas:
            atual = atual.lower()
            esperado = procurado.lower()
        else:
            esperado = procurado
        return atual == esperado if correspondencia_exata else esperado in atual

    if incluir_pasta_inicial and corresponde(pasta_inicial):
        resultado.append(pasta_inicial)

    for dados in listar_subpastas(
        pasta_inicial,
        recursivo=True,
        incluir_objeto=True,
    ):
        pasta = dados["pasta"]
        if corresponde(pasta):
            resultado.append(pasta)

    return resultado


__all__ = [
    "OutlookPastasError",
    "PastaInvalidaError",
    "PastaNaoEncontradaError",
    "PastaCriacaoError",
    "PastaAcaoError",
    "CaminhoPastaError",
    "obter_pasta_padrao",
    "obter_pastas_raiz",
    "contar_subpastas",
    "obter_subpasta_por_indice",
    "obter_subpasta",
    "subpasta_existe",
    "listar_subpastas",
    "obter_pasta_por_caminho",
    "criar_subpasta",
    "criar_caminho_pastas",
    "renomear_pasta",
    "mover_pasta",
    "copiar_pasta",
    "excluir_pasta",
    "exibir_pasta",
    "contar_itens",
    "contar_itens_nao_lidos",
    "listar_itens",
    "obter_dados_pasta",
    "obter_pasta_pai",
    "obter_caminho_pasta",
    "localizar_pastas_por_nome",
]
