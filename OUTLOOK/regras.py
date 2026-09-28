"""
============================================================
Modulo: OUTLOOK / regras.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes reutilizaveis para consultar, criar, configurar, executar,
    habilitar, desabilitar, reordenar e excluir regras do Outlook
    classico por meio do modelo de objetos COM.
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
    from constantes import TipoRegra
    TIPO_REGRA_RECEBIMENTO = int(TipoRegra.RECEBIMENTO)
    TIPO_REGRA_ENVIO = int(TipoRegra.ENVIO)
except ImportError:
    TIPO_REGRA_RECEBIMENTO = 0
    TIPO_REGRA_ENVIO = 1


EXECUTAR_TODAS_MENSAGENS = 0
EXECUTAR_MENSAGENS_LIDAS = 1
EXECUTAR_MENSAGENS_NAO_LIDAS = 2


class OutlookRegrasError(AutomacaoError):
    """Erro base para operacoes relacionadas a regras do Outlook."""


class ColecaoRegrasError(OutlookRegrasError):
    """A colecao Rules nao pode ser obtida, consultada ou salva."""


class RegraNaoEncontradaError(OutlookRegrasError):
    """A regra solicitada nao foi encontrada na colecao."""


class RegraCriacaoError(OutlookRegrasError):
    """A regra nao pode ser criada pelo Outlook."""


class RegraConfiguracaoError(OutlookRegrasError):
    """As condicoes, excecoes ou acoes da regra nao podem ser configuradas."""


class RegraExecucaoError(OutlookRegrasError):
    """A regra nao pode ser executada na pasta solicitada."""


class RegraAcaoError(OutlookRegrasError):
    """A regra nao pode ser habilitada, reordenada ou excluida."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM e retorna um padrao em caso de erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise RegraConfiguracaoError(f"O parametro '{nome}' nao pode ser None.")


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


def _normalizar_lista_textos(
    valores: str | list[str] | tuple[str, ...] | set[str],
    nome: str,
) -> list[str]:
    """Normaliza um texto ou colecao textual, removendo duplicidades."""
    if isinstance(valores, str):
        valores = [valores]
    if not isinstance(valores, (list, tuple, set)) or not valores:
        raise TypeError(f"O parametro '{nome}' deve ser uma colecao nao vazia.")
    resultado: list[str] = []
    for valor in valores:
        texto = _validar_texto(valor, nome)
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def obter_store_padrao(namespace: Any) -> Any:
    """Retorna o DefaultStore da sessao atual do Outlook."""
    _validar_objeto(namespace, "namespace")
    store = _obter_propriedade(namespace, "DefaultStore")
    if store is None:
        raise ColecaoRegrasError("O DefaultStore nao esta disponivel.")
    return store


def obter_colecao_regras(store_ou_namespace: Any) -> Any:
    """Retorna a colecao Rules de um Store ou do DefaultStore do Namespace."""
    _validar_objeto(store_ou_namespace, "store_ou_namespace")
    store = store_ou_namespace
    if _obter_propriedade(store_ou_namespace, "DefaultStore") is not None:
        store = obter_store_padrao(store_ou_namespace)
    try:
        regras = store.GetRules()
    except Exception as error:
        raise ColecaoRegrasError("Nao foi possivel obter a colecao Rules.") from error
    if regras is None:
        raise ColecaoRegrasError("A colecao Rules nao esta disponivel.")
    return regras


def contar_regras(regras_ou_store: Any) -> int:
    """Retorna a quantidade de regras da colecao ou Store informado."""
    _validar_objeto(regras_ou_store, "regras_ou_store")
    regras = regras_ou_store
    if _obter_propriedade(regras_ou_store, "Count") is None:
        regras = obter_colecao_regras(regras_ou_store)
    try:
        return int(regras.Count)
    except Exception as error:
        raise ColecaoRegrasError("Nao foi possivel contar as regras.") from error


def obter_regra_por_indice(regras: Any, indice: int) -> Any:
    """Retorna uma regra pelo indice COM iniciado em um."""
    _validar_objeto(regras, "regras")
    if isinstance(indice, bool) or not isinstance(indice, int):
        raise TypeError("O parametro 'indice' deve ser inteiro.")
    if indice < 1:
        raise ValueError("O parametro 'indice' deve ser maior ou igual a 1.")
    quantidade = contar_regras(regras)
    if indice > quantidade:
        raise RegraNaoEncontradaError(
            f"Nao existe regra no indice {indice}. Quantidade: {quantidade}."
        )
    try:
        return regras.Item(indice)
    except Exception as error:
        raise RegraNaoEncontradaError(
            f"Nao foi possivel obter a regra no indice {indice}."
        ) from error


def obter_regra_por_nome(
    regras: Any,
    nome: str,
    considerar_maiusculas: bool = False,
) -> Any:
    """Localiza uma regra pelo nome de exibicao."""
    procurado = _validar_texto(nome, "nome")
    _validar_booleano(considerar_maiusculas, "considerar_maiusculas")
    for indice in range(1, contar_regras(regras) + 1):
        regra = obter_regra_por_indice(regras, indice)
        atual = str(_obter_propriedade(regra, "Name", "") or "")
        if considerar_maiusculas:
            corresponde = atual == procurado
        else:
            corresponde = atual.lower() == procurado.lower()
        if corresponde:
            return regra
    raise RegraNaoEncontradaError(f"A regra '{procurado}' nao foi encontrada.")


def regra_existe(
    regras: Any,
    nome: str,
    considerar_maiusculas: bool = False,
) -> bool:
    """Verifica se uma regra existe na colecao."""
    try:
        obter_regra_por_nome(regras, nome, considerar_maiusculas)
        return True
    except RegraNaoEncontradaError:
        return False


def obter_dados_regra(regra: Any) -> dict[str, Any]:
    """Converte as principais propriedades de uma regra para dicionario."""
    _validar_objeto(regra, "regra")
    return {
        "nome": str(_obter_propriedade(regra, "Name", "") or ""),
        "habilitada": bool(_obter_propriedade(regra, "Enabled", False)),
        "ordem_execucao": _obter_propriedade(regra, "ExecutionOrder"),
        "tipo": _obter_propriedade(regra, "RuleType"),
        "regra_local": bool(_obter_propriedade(regra, "IsLocalRule", False)),
        "possui_condicoes": _obter_propriedade(regra, "Conditions") is not None,
        "possui_excecoes": _obter_propriedade(regra, "Exceptions") is not None,
        "possui_acoes": _obter_propriedade(regra, "Actions") is not None,
    }


def listar_regras(
    regras_ou_store: Any,
    apenas_habilitadas: bool = False,
    incluir_objeto: bool = True,
) -> list[dict[str, Any]]:
    """Lista regras e seus metadados na ordem de execucao."""
    _validar_booleano(apenas_habilitadas, "apenas_habilitadas")
    _validar_booleano(incluir_objeto, "incluir_objeto")
    regras = regras_ou_store
    if _obter_propriedade(regras_ou_store, "Count") is None:
        regras = obter_colecao_regras(regras_ou_store)
    resultado: list[dict[str, Any]] = []
    for indice in range(1, contar_regras(regras) + 1):
        regra = obter_regra_por_indice(regras, indice)
        if apenas_habilitadas and not bool(_obter_propriedade(regra, "Enabled", False)):
            continue
        dados = obter_dados_regra(regra)
        dados["indice"] = indice
        if incluir_objeto:
            dados["regra"] = regra
        resultado.append(dados)
    return resultado


def criar_regra(
    regras: Any,
    nome: str,
    tipo: int = TIPO_REGRA_RECEBIMENTO,
    retornar_existente: bool = False,
) -> Any:
    """Cria uma regra de recebimento ou envio na colecao Rules."""
    _validar_objeto(regras, "regras")
    nome_validado = _validar_texto(nome, "nome")
    _validar_booleano(retornar_existente, "retornar_existente")
    if isinstance(tipo, bool) or not isinstance(tipo, int):
        raise TypeError("O parametro 'tipo' deve ser inteiro.")
    if tipo not in {TIPO_REGRA_RECEBIMENTO, TIPO_REGRA_ENVIO}:
        raise ValueError("O tipo da regra deve ser recebimento ou envio.")
    if retornar_existente and regra_existe(regras, nome_validado):
        return obter_regra_por_nome(regras, nome_validado)
    try:
        return regras.Create(nome_validado, tipo)
    except Exception as error:
        raise RegraCriacaoError(
            f"Nao foi possivel criar a regra '{nome_validado}'."
        ) from error


def salvar_regras(regras: Any) -> Any:
    """Persiste as alteracoes realizadas na colecao Rules."""
    _validar_objeto(regras, "regras")
    try:
        regras.Save()
        return regras
    except Exception as error:
        raise ColecaoRegrasError("Nao foi possivel salvar as regras.") from error


def definir_regra_habilitada(
    regra: Any,
    habilitada: bool = True,
    regras: Any | None = None,
    salvar: bool = False,
) -> Any:
    """Habilita ou desabilita uma regra e opcionalmente salva a colecao."""
    _validar_objeto(regra, "regra")
    _validar_booleano(habilitada, "habilitada")
    _validar_booleano(salvar, "salvar")
    try:
        regra.Enabled = habilitada
    except Exception as error:
        raise RegraAcaoError("Nao foi possivel alterar o estado da regra.") from error
    if salvar:
        if regras is None:
            raise RegraConfiguracaoError(
                "Informe a colecao 'regras' para persistir a alteracao."
            )
        salvar_regras(regras)
    return regra


def definir_ordem_execucao(
    regra: Any,
    ordem: int,
    regras: Any | None = None,
    salvar: bool = False,
) -> Any:
    """Altera a ordem de execucao da regra."""
    _validar_objeto(regra, "regra")
    if isinstance(ordem, bool) or not isinstance(ordem, int):
        raise TypeError("O parametro 'ordem' deve ser inteiro.")
    if ordem < 1:
        raise ValueError("O parametro 'ordem' deve ser maior ou igual a 1.")
    _validar_booleano(salvar, "salvar")
    try:
        regra.ExecutionOrder = ordem
    except Exception as error:
        raise RegraAcaoError("Nao foi possivel alterar a ordem da regra.") from error
    if salvar:
        if regras is None:
            raise RegraConfiguracaoError(
                "Informe a colecao 'regras' para persistir a alteracao."
            )
        salvar_regras(regras)
    return regra


def renomear_regra(
    regra: Any,
    novo_nome: str,
    regras: Any | None = None,
    salvar: bool = False,
) -> Any:
    """Altera o nome de uma regra."""
    _validar_objeto(regra, "regra")
    nome = _validar_texto(novo_nome, "novo_nome")
    _validar_booleano(salvar, "salvar")
    try:
        regra.Name = nome
    except Exception as error:
        raise RegraAcaoError("Nao foi possivel renomear a regra.") from error
    if salvar:
        if regras is None:
            raise RegraConfiguracaoError(
                "Informe a colecao 'regras' para persistir a alteracao."
            )
        salvar_regras(regras)
    return regra


def executar_regra(
    regra: Any,
    pasta: Any | None = None,
    exibir_progresso: bool = False,
    incluir_subpastas: bool = False,
    opcao_execucao: int = EXECUTAR_TODAS_MENSAGENS,
) -> Any:
    """Executa uma regra uma unica vez com parametros opcionais."""
    _validar_objeto(regra, "regra")
    _validar_booleano(exibir_progresso, "exibir_progresso")
    _validar_booleano(incluir_subpastas, "incluir_subpastas")
    if isinstance(opcao_execucao, bool) or not isinstance(opcao_execucao, int):
        raise TypeError("O parametro 'opcao_execucao' deve ser inteiro.")
    if opcao_execucao not in {
        EXECUTAR_TODAS_MENSAGENS,
        EXECUTAR_MENSAGENS_LIDAS,
        EXECUTAR_MENSAGENS_NAO_LIDAS,
    }:
        raise ValueError("Opcao de execucao de regra invalida.")
    try:
        if pasta is None:
            return regra.Execute(
                exibir_progresso,
                None,
                incluir_subpastas,
                opcao_execucao,
            )
        return regra.Execute(
            exibir_progresso,
            pasta,
            incluir_subpastas,
            opcao_execucao,
        )
    except Exception as error:
        raise RegraExecucaoError("Nao foi possivel executar a regra.") from error


def excluir_regra(
    regras: Any,
    regra_ou_indice: Any,
    salvar: bool = True,
) -> bool:
    """Exclui uma regra por indice, nome ou objeto Rule."""
    _validar_objeto(regras, "regras")
    _validar_booleano(salvar, "salvar")
    if isinstance(regra_ou_indice, bool):
        raise TypeError("O identificador da regra nao pode ser booleano.")
    identificador = regra_ou_indice
    if not isinstance(regra_ou_indice, (int, str)):
        nome = _obter_propriedade(regra_ou_indice, "Name")
        if not nome:
            raise TypeError("Informe indice, nome ou objeto Rule valido.")
        identificador = str(nome)
    try:
        regras.Remove(identificador)
        if salvar:
            salvar_regras(regras)
        return True
    except Exception as error:
        raise RegraAcaoError("Nao foi possivel excluir a regra.") from error


def obter_condicoes(regra: Any) -> Any:
    """Retorna o objeto RuleConditions da regra."""
    _validar_objeto(regra, "regra")
    condicoes = _obter_propriedade(regra, "Conditions")
    if condicoes is None:
        raise RegraConfiguracaoError("As condicoes da regra nao estao disponiveis.")
    return condicoes


def obter_excecoes(regra: Any) -> Any:
    """Retorna o objeto RuleConditions usado como excecoes da regra."""
    _validar_objeto(regra, "regra")
    excecoes = _obter_propriedade(regra, "Exceptions")
    if excecoes is None:
        raise RegraConfiguracaoError("As excecoes da regra nao estao disponiveis.")
    return excecoes


def obter_acoes(regra: Any) -> Any:
    """Retorna o objeto RuleActions da regra."""
    _validar_objeto(regra, "regra")
    acoes = _obter_propriedade(regra, "Actions")
    if acoes is None:
        raise RegraConfiguracaoError("As acoes da regra nao estao disponiveis.")
    return acoes


def definir_condicao_textual(
    condicao: Any,
    valores: str | list[str] | tuple[str, ...] | set[str],
    habilitada: bool = True,
) -> Any:
    """Configura uma condicao textual com Enabled e Text."""
    _validar_objeto(condicao, "condicao")
    textos = _normalizar_lista_textos(valores, "valores")
    _validar_booleano(habilitada, "habilitada")
    try:
        condicao.Enabled = habilitada
        condicao.Text = textos
        return condicao
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar a condicao textual."
        ) from error


def definir_condicao_remetentes(
    condicao: Any,
    remetentes: str | list[str] | tuple[str, ...] | set[str],
    habilitada: bool = True,
    resolver: bool = True,
) -> Any:
    """Configura condicao de remetentes por meio de Recipients."""
    _validar_objeto(condicao, "condicao")
    valores = _normalizar_lista_textos(remetentes, "remetentes")
    _validar_booleano(habilitada, "habilitada")
    _validar_booleano(resolver, "resolver")
    try:
        condicao.Enabled = habilitada
        for valor in valores:
            condicao.Recipients.Add(valor)
        if resolver and not bool(condicao.Recipients.ResolveAll()):
            raise RegraConfiguracaoError(
                "Um ou mais remetentes nao foram resolvidos."
            )
        return condicao
    except RegraConfiguracaoError:
        raise
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar os remetentes da regra."
        ) from error


def definir_acao_mover_para_pasta(
    acao: Any,
    pasta_destino: Any,
    habilitada: bool = True,
) -> Any:
    """Configura uma acao MoveToFolder."""
    _validar_objeto(acao, "acao")
    _validar_objeto(pasta_destino, "pasta_destino")
    _validar_booleano(habilitada, "habilitada")
    try:
        acao.Enabled = habilitada
        acao.Folder = pasta_destino
        return acao
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar a acao de movimentacao."
        ) from error


def definir_acao_categorias(
    acao: Any,
    categorias: str | list[str] | tuple[str, ...] | set[str],
    habilitada: bool = True,
) -> Any:
    """Configura uma acao de atribuicao de categorias."""
    _validar_objeto(acao, "acao")
    valores = _normalizar_lista_textos(categorias, "categorias")
    _validar_booleano(habilitada, "habilitada")
    try:
        acao.Enabled = habilitada
        acao.Categories = valores
        return acao
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar as categorias da regra."
        ) from error


def definir_acao_marcar_importancia(
    acao: Any,
    importancia: int,
    habilitada: bool = True,
) -> Any:
    """Configura uma acao de alteracao de importancia."""
    _validar_objeto(acao, "acao")
    if isinstance(importancia, bool) or not isinstance(importancia, int):
        raise TypeError("O parametro 'importancia' deve ser inteiro.")
    if importancia not in {0, 1, 2}:
        raise ValueError("A importancia deve ser 0, 1 ou 2.")
    _validar_booleano(habilitada, "habilitada")
    try:
        acao.Enabled = habilitada
        acao.Importance = importancia
        return acao
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar a importancia da regra."
        ) from error


def definir_acao_parar_processamento(
    acao: Any,
    habilitada: bool = True,
) -> Any:
    """Configura a acao de parar o processamento das regras seguintes."""
    _validar_objeto(acao, "acao")
    _validar_booleano(habilitada, "habilitada")
    try:
        acao.Enabled = habilitada
        return acao
    except Exception as error:
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar a parada de processamento."
        ) from error


def criar_regra_assunto_mover(
    regras: Any,
    nome: str,
    palavras_assunto: str | list[str] | tuple[str, ...] | set[str],
    pasta_destino: Any,
    habilitada: bool = True,
    parar_processamento: bool = False,
    salvar: bool = True,
) -> Any:
    """Cria uma regra de recebimento para mover mensagens pelo assunto."""
    _validar_booleano(habilitada, "habilitada")
    _validar_booleano(parar_processamento, "parar_processamento")
    _validar_booleano(salvar, "salvar")
    regra = criar_regra(regras, nome, TIPO_REGRA_RECEBIMENTO)
    try:
        definir_condicao_textual(
            obter_condicoes(regra).Subject,
            palavras_assunto,
            habilitada=True,
        )
        definir_acao_mover_para_pasta(
            obter_acoes(regra).MoveToFolder,
            pasta_destino,
            habilitada=True,
        )
        if parar_processamento:
            definir_acao_parar_processamento(
                obter_acoes(regra).StopProcessing,
                habilitada=True,
            )
        regra.Enabled = habilitada
        if salvar:
            salvar_regras(regras)
        return regra
    except Exception as error:
        if isinstance(error, OutlookRegrasError):
            raise
        raise RegraConfiguracaoError(
            "Nao foi possivel configurar a regra de assunto."
        ) from error


def criar_regra_remetente_mover(
    regras: Any,
    nome: str,
    remetentes: str | list[str] | tuple[str, ...] | set[str],
    pasta_destino: Any,
    habilitada: bool = True,
    salvar: bool = True,
) -> Any:
    """Cria uma regra de recebimento para mover mensagens por remetente."""
    _validar_booleano(habilitada, "habilitada")
    _validar_booleano(salvar, "salvar")
    regra = criar_regra(regras, nome, TIPO_REGRA_RECEBIMENTO)
    definir_condicao_remetentes(
        obter_condicoes(regra).From,
        remetentes,
        habilitada=True,
        resolver=True,
    )
    definir_acao_mover_para_pasta(
        obter_acoes(regra).MoveToFolder,
        pasta_destino,
        habilitada=True,
    )
    regra.Enabled = habilitada
    if salvar:
        salvar_regras(regras)
    return regra


__all__ = [
    "TIPO_REGRA_RECEBIMENTO",
    "TIPO_REGRA_ENVIO",
    "EXECUTAR_TODAS_MENSAGENS",
    "EXECUTAR_MENSAGENS_LIDAS",
    "EXECUTAR_MENSAGENS_NAO_LIDAS",
    "OutlookRegrasError",
    "ColecaoRegrasError",
    "RegraNaoEncontradaError",
    "RegraCriacaoError",
    "RegraConfiguracaoError",
    "RegraExecucaoError",
    "RegraAcaoError",
    "obter_store_padrao",
    "obter_colecao_regras",
    "contar_regras",
    "obter_regra_por_indice",
    "obter_regra_por_nome",
    "regra_existe",
    "obter_dados_regra",
    "listar_regras",
    "criar_regra",
    "salvar_regras",
    "definir_regra_habilitada",
    "definir_ordem_execucao",
    "renomear_regra",
    "executar_regra",
    "excluir_regra",
    "obter_condicoes",
    "obter_excecoes",
    "obter_acoes",
    "definir_condicao_textual",
    "definir_condicao_remetentes",
    "definir_acao_mover_para_pasta",
    "definir_acao_categorias",
    "definir_acao_marcar_importancia",
    "definir_acao_parar_processamento",
    "criar_regra_assunto_mover",
    "criar_regra_remetente_mover",
]
