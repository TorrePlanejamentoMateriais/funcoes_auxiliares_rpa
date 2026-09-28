"""
============================================================
Modulo: OUTLOOK / pesquisa.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes reutilizaveis para pesquisar e filtrar itens em pastas do
    Outlook classico. Oferece pesquisa local em colecoes COM, Restrict,
    Find/FindNext e construcao segura de filtros Outlook.
Dependencias:
    - Outlook classico para Windows
    - excecoes.py
============================================================
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Callable, Iterable

try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base usada quando o modulo central nao esta disponivel."""


FORMATO_DATA_FILTRO = "%m/%d/%Y %I:%M %p"


class OutlookPesquisaError(AutomacaoError):
    """Erro base para pesquisas e filtros no Outlook."""


class FiltroPesquisaError(OutlookPesquisaError):
    """O filtro informado possui sintaxe ou dados invalidos."""


class ColecaoPesquisaError(OutlookPesquisaError):
    """A colecao Items nao pode ser consultada ou percorrida."""


class ItemNaoEncontradoError(OutlookPesquisaError):
    """Nenhum item correspondente foi encontrado."""


class CampoPesquisaError(OutlookPesquisaError):
    """O campo solicitado nao esta disponivel para pesquisa."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM e retorna um padrao em caso de erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise ColecaoPesquisaError(f"O parametro '{nome}' nao pode ser None.")


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


def _normalizar_campo(campo: str) -> str:
    """Normaliza um campo para a sintaxe [Campo] do Outlook."""
    texto = _validar_texto(campo, "campo")
    if texto.startswith("[") and texto.endswith("]"):
        return texto
    return f"[{texto}]"


def escapar_texto_filtro(valor: str) -> str:
    """Escapa apostrofos para filtros Find e Restrict."""
    texto = _validar_texto(valor, "valor", permitir_vazio=True)
    return texto.replace("'", "''")


def formatar_data_filtro(valor: date | datetime) -> str:
    """Formata date ou datetime segundo o formato esperado pelo Outlook."""
    if isinstance(valor, datetime):
        momento = valor
    elif isinstance(valor, date):
        momento = datetime.combine(valor, datetime.min.time())
    else:
        raise TypeError("O parametro 'valor' deve ser date ou datetime.")
    return momento.strftime(FORMATO_DATA_FILTRO)


def criar_filtro_texto(
    campo: str,
    valor: str,
    operador: str = "=",
) -> str:
    """Cria um filtro textual simples para Find ou Restrict."""
    campo_formatado = _normalizar_campo(campo)
    texto = escapar_texto_filtro(valor)
    op = _validar_texto(operador, "operador")
    if op not in {"=", "<>", ">", ">=", "<", "<="}:
        raise FiltroPesquisaError(f"Operador nao suportado: '{operador}'.")
    return f"{campo_formatado} {op} '{texto}'"


def criar_filtro_numero(campo: str, valor: int | float, operador: str = "=") -> str:
    """Cria um filtro numerico simples."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise TypeError("O parametro 'valor' deve ser numerico.")
    op = _validar_texto(operador, "operador")
    if op not in {"=", "<>", ">", ">=", "<", "<="}:
        raise FiltroPesquisaError(f"Operador nao suportado: '{operador}'.")
    return f"{_normalizar_campo(campo)} {op} {valor}"


def criar_filtro_booleano(campo: str, valor: bool) -> str:
    """Cria um filtro booleano Outlook com True ou False."""
    _validar_booleano(valor, "valor")
    return f"{_normalizar_campo(campo)} = {'True' if valor else 'False'}"


def criar_filtro_data(
    campo: str,
    valor: date | datetime,
    operador: str = ">=",
) -> str:
    """Cria um filtro de data para Find ou Restrict."""
    op = _validar_texto(operador, "operador")
    if op not in {"=", "<>", ">", ">=", "<", "<="}:
        raise FiltroPesquisaError(f"Operador nao suportado: '{operador}'.")
    return f"{_normalizar_campo(campo)} {op} '{formatar_data_filtro(valor)}'"


def combinar_filtros(
    *filtros: str,
    operador: str = "AND",
    adicionar_parenteses: bool = True,
) -> str:
    """Combina dois ou mais filtros com AND ou OR."""
    if len(filtros) < 1:
        raise FiltroPesquisaError("Informe ao menos um filtro.")
    op = _validar_texto(operador, "operador").upper()
    if op not in {"AND", "OR"}:
        raise FiltroPesquisaError("O operador deve ser AND ou OR.")
    _validar_booleano(adicionar_parenteses, "adicionar_parenteses")
    normalizados = [_validar_texto(filtro, "filtro") for filtro in filtros]
    if adicionar_parenteses:
        normalizados = [f"({filtro})" for filtro in normalizados]
    return f" {op} ".join(normalizados)


def criar_filtro_intervalo_datas(
    campo: str,
    inicio: date | datetime,
    fim: date | datetime,
    inclusivo: bool = True,
) -> str:
    """Cria filtro de data inicial e final para um mesmo campo."""
    _validar_booleano(inclusivo, "inclusivo")
    inicio_dt = inicio if isinstance(inicio, datetime) else datetime.combine(inicio, datetime.min.time()) if isinstance(inicio, date) else None
    fim_dt = fim if isinstance(fim, datetime) else datetime.combine(fim, datetime.min.time()) if isinstance(fim, date) else None
    if inicio_dt is None or fim_dt is None:
        raise TypeError("Inicio e fim devem ser date ou datetime.")
    if fim_dt < inicio_dt:
        raise FiltroPesquisaError("A data final nao pode ser anterior a inicial.")
    filtro_inicio = criar_filtro_data(campo, inicio_dt, ">=" if inclusivo else ">")
    filtro_fim = criar_filtro_data(campo, fim_dt, "<=" if inclusivo else "<")
    return combinar_filtros(filtro_inicio, filtro_fim)


def obter_colecao_itens(pasta_ou_itens: Any) -> Any:
    """Retorna uma colecao Items a partir de uma pasta ou da propria colecao."""
    _validar_objeto(pasta_ou_itens, "pasta_ou_itens")
    itens = _obter_propriedade(pasta_ou_itens, "Items")
    if itens is not None:
        return itens
    if _obter_propriedade(pasta_ou_itens, "Count") is not None:
        return pasta_ou_itens
    raise ColecaoPesquisaError("O objeto informado nao possui uma colecao Items valida.")


def configurar_colecao(
    itens: Any,
    ordenar_por: str | None = None,
    ordem_decrescente: bool = False,
    incluir_recorrencias: bool | None = None,
) -> Any:
    """Configura ordenacao e recorrencias de uma colecao Items."""
    _validar_objeto(itens, "itens")
    _validar_booleano(ordem_decrescente, "ordem_decrescente")
    try:
        if ordenar_por is not None:
            itens.Sort(_normalizar_campo(ordenar_por), ordem_decrescente)
        if incluir_recorrencias is not None:
            _validar_booleano(incluir_recorrencias, "incluir_recorrencias")
            itens.IncludeRecurrences = incluir_recorrencias
        return itens
    except Exception as error:
        raise ColecaoPesquisaError("Nao foi possivel configurar a colecao Items.") from error


def colecao_para_lista(itens: Any, limite: int | None = None) -> list[Any]:
    """Converte uma colecao COM iniciada em um para lista Python."""
    _validar_objeto(itens, "itens")
    if limite is not None:
        if isinstance(limite, bool) or not isinstance(limite, int):
            raise TypeError("O parametro 'limite' deve ser inteiro ou None.")
        if limite < 1:
            raise ValueError("O parametro 'limite' deve ser maior que zero.")
    try:
        quantidade = int(itens.Count)
    except Exception as error:
        raise ColecaoPesquisaError("Nao foi possivel contar os itens.") from error
    quantidade = quantidade if limite is None else min(quantidade, limite)
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            resultado.append(itens.Item(indice))
        except Exception as error:
            raise ColecaoPesquisaError(
                f"Nao foi possivel obter o item no indice {indice}."
            ) from error
    return resultado


def restringir_itens(
    pasta_ou_itens: Any,
    filtro: str,
    ordenar_por: str | None = None,
    ordem_decrescente: bool = False,
    incluir_recorrencias: bool | None = None,
    limite: int | None = None,
) -> list[Any]:
    """Aplica Items.Restrict e converte o resultado para uma lista."""
    expressao = _validar_texto(filtro, "filtro")
    itens = obter_colecao_itens(pasta_ou_itens)
    itens = configurar_colecao(
        itens,
        ordenar_por=ordenar_por,
        ordem_decrescente=ordem_decrescente,
        incluir_recorrencias=incluir_recorrencias,
    )
    try:
        filtrados = itens.Restrict(expressao)
    except Exception as error:
        raise FiltroPesquisaError(
            f"Nao foi possivel aplicar o filtro: {expressao}"
        ) from error
    return colecao_para_lista(filtrados, limite=limite)


def localizar_primeiro_item(pasta_ou_itens: Any, filtro: str) -> Any:
    """Localiza o primeiro item utilizando Items.Find."""
    expressao = _validar_texto(filtro, "filtro")
    itens = obter_colecao_itens(pasta_ou_itens)
    try:
        item = itens.Find(expressao)
    except Exception as error:
        raise FiltroPesquisaError(
            f"Nao foi possivel executar Find com o filtro: {expressao}"
        ) from error
    if item is None:
        raise ItemNaoEncontradoError("Nenhum item corresponde ao filtro informado.")
    return item


def localizar_itens_find(
    pasta_ou_itens: Any,
    filtro: str,
    limite: int | None = None,
) -> list[Any]:
    """Percorre resultados utilizando Find e FindNext."""
    expressao = _validar_texto(filtro, "filtro")
    itens = obter_colecao_itens(pasta_ou_itens)
    if limite is not None:
        if isinstance(limite, bool) or not isinstance(limite, int):
            raise TypeError("O parametro 'limite' deve ser inteiro ou None.")
        if limite < 1:
            raise ValueError("O parametro 'limite' deve ser maior que zero.")
    try:
        item = itens.Find(expressao)
        resultado = []
        while item is not None:
            resultado.append(item)
            if limite is not None and len(resultado) >= limite:
                break
            item = itens.FindNext()
        return resultado
    except Exception as error:
        raise FiltroPesquisaError("Nao foi possivel percorrer Find/FindNext.") from error


def filtrar_itens_python(
    pasta_ou_itens: Any,
    predicado: Callable[[Any], bool],
    limite: int | None = None,
) -> list[Any]:
    """Filtra itens em Python utilizando uma funcao predicado."""
    if not callable(predicado):
        raise TypeError("O parametro 'predicado' deve ser chamavel.")
    itens = colecao_para_lista(obter_colecao_itens(pasta_ou_itens))
    resultado = []
    for item in itens:
        try:
            corresponde = bool(predicado(item))
        except Exception as error:
            raise OutlookPesquisaError("O predicado gerou erro ao avaliar um item.") from error
        if corresponde:
            resultado.append(item)
            if limite is not None and len(resultado) >= limite:
                break
    return resultado


def buscar_por_propriedade(
    pasta_ou_itens: Any,
    propriedade: str,
    valor: Any,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
    limite: int | None = None,
) -> list[Any]:
    """Pesquisa localmente por uma propriedade COM simples."""
    nome = _validar_texto(propriedade, "propriedade")
    _validar_booleano(correspondencia_exata, "correspondencia_exata")
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")

    def predicado(item: Any) -> bool:
        atual = _obter_propriedade(item, nome, None)
        if atual is None:
            return False
        if isinstance(atual, str) and isinstance(valor, str):
            a = atual if considerar_maiusculas else atual.lower()
            b = valor if considerar_maiusculas else valor.lower()
            return a == b if correspondencia_exata else b in a
        return atual == valor

    return filtrar_itens_python(pasta_ou_itens, predicado, limite=limite)


def buscar_emails(
    pasta_ou_itens: Any,
    assunto: str | None = None,
    remetente: str | None = None,
    nao_lido: bool | None = None,
    recebido_inicio: datetime | None = None,
    recebido_fim: datetime | None = None,
    limite: int | None = None,
) -> list[Any]:
    """Pesquisa e-mails por assunto, remetente, leitura e intervalo de recebimento."""
    if assunto is not None:
        assunto = _validar_texto(assunto, "assunto")
    if remetente is not None:
        remetente = _validar_texto(remetente, "remetente")
    if nao_lido is not None:
        _validar_booleano(nao_lido, "nao_lido")
    for nome, valor in {"recebido_inicio": recebido_inicio, "recebido_fim": recebido_fim}.items():
        if valor is not None and not isinstance(valor, datetime):
            raise TypeError(f"O parametro '{nome}' deve ser datetime ou None.")
    if recebido_inicio and recebido_fim and recebido_fim < recebido_inicio:
        raise FiltroPesquisaError("O fim do intervalo nao pode ser anterior ao inicio.")
    if all(valor is None for valor in (assunto, remetente, nao_lido, recebido_inicio, recebido_fim)):
        raise FiltroPesquisaError("Informe ao menos um criterio de pesquisa.")

    def predicado(item: Any) -> bool:
        if assunto is not None:
            atual = str(_obter_propriedade(item, "Subject", "") or "")
            if assunto.lower() not in atual.lower():
                return False
        if remetente is not None:
            nome = str(_obter_propriedade(item, "SenderName", "") or "")
            email = str(_obter_propriedade(item, "SenderEmailAddress", "") or "")
            if remetente.lower() not in nome.lower() and remetente.lower() not in email.lower():
                return False
        if nao_lido is not None and bool(_obter_propriedade(item, "UnRead", False)) != nao_lido:
            return False
        recebido = _obter_propriedade(item, "ReceivedTime")
        if recebido_inicio is not None and (not isinstance(recebido, datetime) or recebido < recebido_inicio):
            return False
        if recebido_fim is not None and (not isinstance(recebido, datetime) or recebido > recebido_fim):
            return False
        return True

    return filtrar_itens_python(pasta_ou_itens, predicado, limite=limite)


def buscar_por_entry_id(namespace: Any, entry_id: str, store_id: str | None = None) -> Any:
    """Localiza um item diretamente pelo EntryID e StoreID opcional."""
    _validar_objeto(namespace, "namespace")
    identificador = _validar_texto(entry_id, "entry_id")
    if store_id is not None:
        store_id = _validar_texto(store_id, "store_id")
    try:
        item = (
            namespace.GetItemFromID(identificador, store_id)
            if store_id is not None
            else namespace.GetItemFromID(identificador)
        )
    except Exception as error:
        raise ItemNaoEncontradoError(
            f"Nao foi possivel localizar o item '{identificador}'."
        ) from error
    if item is None:
        raise ItemNaoEncontradoError(f"O item '{identificador}' nao foi encontrado.")
    return item


def primeiro_ou_none(itens: Iterable[Any]) -> Any | None:
    """Retorna o primeiro item de um iteravel ou None."""
    try:
        return next(iter(itens))
    except StopIteration:
        return None
    except TypeError as error:
        raise TypeError("O parametro 'itens' deve ser iteravel.") from error


def exigir_primeiro(itens: Iterable[Any], mensagem: str = "Nenhum item encontrado.") -> Any:
    """Retorna o primeiro item ou gera ItemNaoEncontradoError."""
    item = primeiro_ou_none(itens)
    if item is None:
        raise ItemNaoEncontradoError(_validar_texto(mensagem, "mensagem"))
    return item


__all__ = [
    "FORMATO_DATA_FILTRO",
    "OutlookPesquisaError",
    "FiltroPesquisaError",
    "ColecaoPesquisaError",
    "ItemNaoEncontradoError",
    "CampoPesquisaError",
    "escapar_texto_filtro",
    "formatar_data_filtro",
    "criar_filtro_texto",
    "criar_filtro_numero",
    "criar_filtro_booleano",
    "criar_filtro_data",
    "combinar_filtros",
    "criar_filtro_intervalo_datas",
    "obter_colecao_itens",
    "configurar_colecao",
    "colecao_para_lista",
    "restringir_itens",
    "localizar_primeiro_item",
    "localizar_itens_find",
    "filtrar_itens_python",
    "buscar_por_propriedade",
    "buscar_emails",
    "buscar_por_entry_id",
    "primeiro_ou_none",
    "exigir_primeiro",
]
