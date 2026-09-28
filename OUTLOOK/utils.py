"""
============================================================
Modulo: OUTLOOK / utils.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes utilitarias compartilhadas pelos modulos da biblioteca de
    automacao do Outlook classico. Centraliza validacoes, normalizacao,
    acesso seguro a propriedades COM, manipulacao de colecoes, datas,
    caminhos e nomes de arquivos.
Dependencias:
    - Nenhuma dependencia externa
============================================================
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable, Mapping
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, TypeVar


T = TypeVar("T")

CARACTERES_INVALIDOS_ARQUIVO = '<>:"/\\|?*'
FORMATO_DATA_OUTLOOK = "%m/%d/%Y %I:%M %p"


class OutlookUtilsError(Exception):
    """Erro base das funcoes utilitarias do Outlook."""


class ValidacaoUtilError(OutlookUtilsError, ValueError):
    """Um valor informado nao atende aos criterios de validacao."""


class PropriedadeCOMError(OutlookUtilsError):
    """Uma propriedade COM nao pode ser obtida ou definida."""


class ColecaoCOMError(OutlookUtilsError):
    """Uma colecao COM nao pode ser consultada ou convertida."""


class CaminhoUtilError(OutlookUtilsError):
    """Um caminho ou nome de arquivo e invalido."""


def obter_propriedade(
    objeto: Any,
    nome: str,
    padrao: Any = None,
    gerar_erro: bool = False,
) -> Any:
    """Obtem uma propriedade de forma segura.

    Args:
        objeto: Objeto Python ou COM.
        nome: Nome da propriedade.
        padrao: Valor retornado quando a propriedade falha.
        gerar_erro: Se True, converte a falha em PropriedadeCOMError.
    """
    validar_booleano(gerar_erro, "gerar_erro")
    nome = validar_texto(nome, "nome")
    if objeto is None:
        if gerar_erro:
            raise PropriedadeCOMError("O objeto informado e None.")
        return padrao
    try:
        return getattr(objeto, nome)
    except Exception as error:
        if gerar_erro:
            raise PropriedadeCOMError(
                f"Nao foi possivel obter a propriedade '{nome}'."
            ) from error
        return padrao


def definir_propriedade(objeto: Any, nome: str, valor: Any) -> Any:
    """Define uma propriedade e retorna o proprio objeto."""
    validar_objeto(objeto, "objeto")
    nome = validar_texto(nome, "nome")
    try:
        setattr(objeto, nome, valor)
        return objeto
    except Exception as error:
        raise PropriedadeCOMError(
            f"Nao foi possivel definir a propriedade '{nome}'."
        ) from error


def propriedade_existe(objeto: Any, nome: str) -> bool:
    """Verifica se uma propriedade pode ser acessada sem erro."""
    if objeto is None:
        return False
    nome = validar_texto(nome, "nome")
    try:
        getattr(objeto, nome)
        return True
    except Exception:
        return False


def validar_objeto(objeto: Any, nome: str = "objeto") -> Any:
    """Garante que um objeto obrigatorio nao seja None."""
    nome = validar_texto(nome, "nome")
    if objeto is None:
        raise ValidacaoUtilError(f"O parametro '{nome}' nao pode ser None.")
    return objeto


def validar_texto(
    valor: str,
    nome: str = "valor",
    permitir_vazio: bool = False,
    remover_espacos: bool = True,
) -> str:
    """Valida texto com controle sobre vazio e remocao de espacos."""
    validar_booleano(permitir_vazio, "permitir_vazio")
    validar_booleano(remover_espacos, "remover_espacos")
    if not isinstance(valor, str):
        raise TypeError(f"O parametro '{nome}' deve ser uma string.")
    texto = valor.strip() if remover_espacos else valor
    if not texto and not permitir_vazio:
        raise ValidacaoUtilError(f"O parametro '{nome}' nao pode estar vazio.")
    return texto


def validar_booleano(valor: bool, nome: str = "valor") -> bool:
    """Valida estritamente um valor booleano."""
    if not isinstance(valor, bool):
        raise TypeError(f"O parametro '{nome}' deve ser booleano.")
    return valor


def validar_inteiro(
    valor: int,
    nome: str = "valor",
    minimo: int | None = None,
    maximo: int | None = None,
) -> int:
    """Valida um inteiro e limites opcionais."""
    if isinstance(valor, bool) or not isinstance(valor, int):
        raise TypeError(f"O parametro '{nome}' deve ser inteiro.")
    if minimo is not None and valor < minimo:
        raise ValidacaoUtilError(
            f"O parametro '{nome}' deve ser maior ou igual a {minimo}."
        )
    if maximo is not None and valor > maximo:
        raise ValidacaoUtilError(
            f"O parametro '{nome}' deve ser menor ou igual a {maximo}."
        )
    return valor


def validar_numero(
    valor: int | float,
    nome: str = "valor",
    minimo: int | float | None = None,
    maximo: int | float | None = None,
) -> int | float:
    """Valida um numero inteiro ou decimal."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        raise TypeError(f"O parametro '{nome}' deve ser numerico.")
    if minimo is not None and valor < minimo:
        raise ValidacaoUtilError(
            f"O parametro '{nome}' deve ser maior ou igual a {minimo}."
        )
    if maximo is not None and valor > maximo:
        raise ValidacaoUtilError(
            f"O parametro '{nome}' deve ser menor ou igual a {maximo}."
        )
    return valor


def validar_data(
    valor: date | datetime | None,
    nome: str = "data",
    permitir_none: bool = False,
) -> date | datetime | None:
    """Valida valores date, datetime e None opcional."""
    validar_booleano(permitir_none, "permitir_none")
    if valor is None and permitir_none:
        return None
    if not isinstance(valor, (date, datetime)):
        raise TypeError(
            f"O parametro '{nome}' deve ser date ou datetime."
        )
    return valor


def normalizar_email(email: str, permitir_vazio: bool = False) -> str:
    """Normaliza e valida estruturalmente um endereco de e-mail."""
    texto = validar_texto(
        email,
        "email",
        permitir_vazio=permitir_vazio,
    )
    if not texto:
        return ""
    if texto.count("@") != 1:
        raise ValidacaoUtilError(f"Endereco de e-mail invalido: '{email}'.")
    usuario, dominio = texto.split("@", 1)
    if not usuario or not dominio or "." not in dominio:
        raise ValidacaoUtilError(f"Endereco de e-mail invalido: '{email}'.")
    return texto.lower()


def normalizar_lista_textos(
    valores: str | Iterable[str],
    nome: str = "valores",
    separadores: tuple[str, ...] = (";", ","),
    remover_duplicados: bool = True,
) -> list[str]:
    """Converte texto ou iteravel em lista textual normalizada."""
    validar_booleano(remover_duplicados, "remover_duplicados")
    if isinstance(valores, str):
        texto = valores
        for separador in separadores[1:]:
            texto = texto.replace(separador, separadores[0])
        valores = texto.split(separadores[0])
    elif not isinstance(valores, Iterable):
        raise TypeError(f"O parametro '{nome}' deve ser texto ou iteravel.")

    resultado: list[str] = []
    for valor in valores:
        texto = validar_texto(valor, nome, permitir_vazio=True)
        if not texto:
            continue
        if not remover_duplicados or texto not in resultado:
            resultado.append(texto)
    if not resultado:
        raise ValidacaoUtilError(f"O parametro '{nome}' nao possui valores validos.")
    return resultado


def normalizar_emails(
    emails: str | Iterable[str],
    remover_duplicados: bool = True,
) -> list[str]:
    """Normaliza uma colecao de enderecos de e-mail."""
    valores = normalizar_lista_textos(
        emails,
        nome="emails",
        remover_duplicados=False,
    )
    resultado: list[str] = []
    for email in valores:
        normalizado = normalizar_email(email)
        if not remover_duplicados or normalizado not in resultado:
            resultado.append(normalizado)
    return resultado


def comparar_textos(
    atual: str,
    esperado: str,
    correspondencia_exata: bool = True,
    considerar_maiusculas: bool = False,
) -> bool:
    """Compara textos de forma exata ou parcial."""
    validar_booleano(correspondencia_exata, "correspondencia_exata")
    validar_booleano(considerar_maiusculas, "considerar_maiusculas")
    atual = validar_texto(atual, "atual", permitir_vazio=True)
    esperado = validar_texto(esperado, "esperado", permitir_vazio=True)
    if not considerar_maiusculas:
        atual = atual.casefold()
        esperado = esperado.casefold()
    return atual == esperado if correspondencia_exata else esperado in atual


def escapar_filtro_outlook(valor: str) -> str:
    """Escapa apostrofos utilizados em filtros Find e Restrict."""
    return validar_texto(
        valor,
        "valor",
        permitir_vazio=True,
        remover_espacos=False,
    ).replace("'", "''")


def formatar_data_outlook(valor: date | datetime) -> str:
    """Formata data para filtros Find e Restrict do Outlook."""
    validar_data(valor)
    momento = (
        valor
        if isinstance(valor, datetime)
        else datetime.combine(valor, datetime.min.time())
    )
    return momento.strftime(FORMATO_DATA_OUTLOOK)


def contar_colecao(colecao: Any) -> int:
    """Retorna Count de uma colecao COM."""
    validar_objeto(colecao, "colecao")
    try:
        return int(colecao.Count)
    except Exception as error:
        raise ColecaoCOMError("Nao foi possivel contar a colecao COM.") from error


def obter_item_colecao(colecao: Any, indice: int) -> Any:
    """Obtem um item de colecao COM pelo indice iniciado em um."""
    validar_objeto(colecao, "colecao")
    validar_inteiro(indice, "indice", minimo=1)
    quantidade = contar_colecao(colecao)
    if indice > quantidade:
        raise ColecaoCOMError(
            f"Indice {indice} fora da colecao. Quantidade: {quantidade}."
        )
    try:
        return colecao.Item(indice)
    except Exception as error:
        raise ColecaoCOMError(
            f"Nao foi possivel obter o item no indice {indice}."
        ) from error


def colecao_para_lista(
    colecao: Any,
    limite: int | None = None,
    transformacao: Callable[[Any], T] | None = None,
    ignorar_erros: bool = False,
) -> list[Any] | list[T]:
    """Converte uma colecao COM iniciada em um para lista Python."""
    validar_objeto(colecao, "colecao")
    validar_booleano(ignorar_erros, "ignorar_erros")
    if limite is not None:
        validar_inteiro(limite, "limite", minimo=1)
    if transformacao is not None and not callable(transformacao):
        raise TypeError("O parametro 'transformacao' deve ser chamavel ou None.")

    quantidade = contar_colecao(colecao)
    quantidade = quantidade if limite is None else min(quantidade, limite)
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            item = colecao.Item(indice)
            resultado.append(
                transformacao(item) if transformacao is not None else item
            )
        except Exception as error:
            if ignorar_erros:
                continue
            raise ColecaoCOMError(
                f"Nao foi possivel processar o item no indice {indice}."
            ) from error
    return resultado


def primeiro_ou_none(valores: Iterable[T]) -> T | None:
    """Retorna o primeiro item de um iteravel ou None."""
    try:
        return next(iter(valores))
    except StopIteration:
        return None
    except TypeError as error:
        raise TypeError("O parametro 'valores' deve ser iteravel.") from error


def remover_acentos(texto: str) -> str:
    """Remove marcas de acentuacao de um texto Unicode."""
    valor = validar_texto(texto, "texto", permitir_vazio=True)
    normalizado = unicodedata.normalize("NFKD", valor)
    return "".join(
        caractere
        for caractere in normalizado
        if not unicodedata.combining(caractere)
    )


def sanitizar_nome_arquivo(
    nome: str,
    substituto: str = "_",
    tamanho_maximo: int = 180,
) -> str:
    """Remove caracteres invalidos de um nome de arquivo Windows."""
    nome = validar_texto(nome, "nome")
    if not isinstance(substituto, str):
        raise TypeError("O parametro 'substituto' deve ser uma string.")
    validar_inteiro(tamanho_maximo, "tamanho_maximo", minimo=1)
    traducao = str.maketrans({caractere: substituto for caractere in CARACTERES_INVALIDOS_ARQUIVO})
    resultado = nome.translate(traducao)
    resultado = re.sub(r"\s+", " ", resultado).strip(" .")
    resultado = resultado[:tamanho_maximo].rstrip(" .")
    if not resultado:
        raise CaminhoUtilError("O nome do arquivo ficou vazio apos a sanitizacao.")
    return resultado


def validar_caminho_arquivo(
    caminho: str | Path,
    deve_existir: bool = True,
    extensoes: Iterable[str] | None = None,
) -> Path:
    """Valida um caminho de arquivo e extensoes permitidas."""
    validar_booleano(deve_existir, "deve_existir")
    if not isinstance(caminho, (str, Path)):
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    arquivo = Path(caminho).expanduser().resolve()
    if deve_existir and not arquivo.is_file():
        raise FileNotFoundError(f"O arquivo nao existe: {arquivo}")
    if extensoes is not None:
        permitidas = {
            extensao.lower() if extensao.startswith(".") else f".{extensao.lower()}"
            for extensao in normalizar_lista_textos(list(extensoes), "extensoes")
        }
        if arquivo.suffix.lower() not in permitidas:
            raise CaminhoUtilError(
                f"Extensao '{arquivo.suffix}' nao permitida: {sorted(permitidas)}."
            )
    return arquivo


def garantir_diretorio(caminho: str | Path) -> Path:
    """Cria um diretorio, incluindo os niveis pais, e retorna seu Path."""
    if not isinstance(caminho, (str, Path)):
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    diretorio = Path(caminho).expanduser().resolve()
    try:
        diretorio.mkdir(parents=True, exist_ok=True)
        return diretorio
    except OSError as error:
        raise CaminhoUtilError(
            f"Nao foi possivel criar o diretorio: {diretorio}"
        ) from error


def caminho_unico(caminho: str | Path) -> Path:
    """Gera um caminho nao existente adicionando sufixo numerico."""
    if not isinstance(caminho, (str, Path)):
        raise TypeError("O parametro 'caminho' deve ser string ou Path.")
    original = Path(caminho).expanduser().resolve()
    if not original.exists():
        return original
    contador = 1
    while True:
        candidato = original.with_name(
            f"{original.stem}_{contador}{original.suffix}"
        )
        if not candidato.exists():
            return candidato
        contador += 1


def mesclar_dicionarios(
    *dicionarios: Mapping[str, Any],
    ignorar_none: bool = False,
) -> dict[str, Any]:
    """Mescla mapeamentos da esquerda para a direita."""
    validar_booleano(ignorar_none, "ignorar_none")
    resultado: dict[str, Any] = {}
    for dicionario in dicionarios:
        if not isinstance(dicionario, Mapping):
            raise TypeError("Todos os valores devem ser mapeamentos.")
        for chave, valor in dicionario.items():
            if ignorar_none and valor is None:
                continue
            resultado[chave] = valor
    return resultado


def mascarar_texto(
    texto: str,
    visiveis_inicio: int = 2,
    visiveis_fim: int = 2,
    caractere: str = "*",
) -> str:
    """Mascara parte de um texto para uso seguro em logs."""
    texto = validar_texto(
        texto,
        "texto",
        permitir_vazio=True,
        remover_espacos=False,
    )
    validar_inteiro(visiveis_inicio, "visiveis_inicio", minimo=0)
    validar_inteiro(visiveis_fim, "visiveis_fim", minimo=0)
    if not isinstance(caractere, str) or len(caractere) != 1:
        raise TypeError("O parametro 'caractere' deve possuir um unico caractere.")
    visiveis = visiveis_inicio + visiveis_fim
    if len(texto) <= visiveis:
        return caractere * len(texto)
    meio = caractere * (len(texto) - visiveis)
    final = texto[-visiveis_fim:] if visiveis_fim else ""
    return f"{texto[:visiveis_inicio]}{meio}{final}"


def liberar_objeto_com(objeto: Any) -> None:
    """Remove uma referencia local a objeto COM sem forcar encerramento."""
    if objeto is None:
        return
    try:
        del objeto
    except Exception:
        return


__all__ = [
    "CARACTERES_INVALIDOS_ARQUIVO",
    "FORMATO_DATA_OUTLOOK",
    "OutlookUtilsError",
    "ValidacaoUtilError",
    "PropriedadeCOMError",
    "ColecaoCOMError",
    "CaminhoUtilError",
    "obter_propriedade",
    "definir_propriedade",
    "propriedade_existe",
    "validar_objeto",
    "validar_texto",
    "validar_booleano",
    "validar_inteiro",
    "validar_numero",
    "validar_data",
    "normalizar_email",
    "normalizar_lista_textos",
    "normalizar_emails",
    "comparar_textos",
    "escapar_filtro_outlook",
    "formatar_data_outlook",
    "contar_colecao",
    "obter_item_colecao",
    "colecao_para_lista",
    "primeiro_ou_none",
    "remover_acentos",
    "sanitizar_nome_arquivo",
    "validar_caminho_arquivo",
    "garantir_diretorio",
    "caminho_unico",
    "mesclar_dicionarios",
    "mascarar_texto",
    "liberar_objeto_com",
]
