"""
============================================================
Modulo: OUTLOOK / modelos.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Modelos de dados reutilizaveis para a biblioteca de automacao do
    Outlook classico. As classes representam anexos, contas, contatos,
    destinatarios, e-mails, compromissos, reunioes e resultados de
    operacoes, sem depender diretamente de objetos COM.
Dependencias:
    - Nenhuma dependencia externa
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial dos modelos da biblioteca Outlook.
        - Inclusao de validacao, serializacao e conversao de dados.
============================================================
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field, fields, replace
from datetime import date, datetime
from enum import IntEnum
from pathlib import Path
from typing import Any, ClassVar, Mapping, TypeVar


TModelo = TypeVar("TModelo", bound="ModeloBase")


class ModeloValidacaoError(ValueError):
    """Erro gerado quando os dados de um modelo sao invalidos."""


class StatusOperacao(IntEnum):
    """Status numerico de uma operacao executada pela automacao."""

    PENDENTE = 0
    SUCESSO = 1
    ERRO = 2
    IGNORADO = 3
    PARCIAL = 4


class TipoDestinatarioModelo(IntEnum):
    """Tipos de destinatario utilizados por e-mails e reunioes."""

    ORGANIZADOR = 0
    PARA = 1
    CC = 2
    CCO = 3

    OBRIGATORIO = 1
    OPCIONAL = 2
    RECURSO = 3


@dataclass(slots=True)
class ModeloBase:
    """Classe base com validacao, copia e serializacao para dicionario."""

    CAMPOS_SENSIVEIS: ClassVar[set[str]] = set()

    def __post_init__(self) -> None:
        """Executa a validacao apos a inicializacao do dataclass."""
        self.validar()

    def validar(self) -> None:
        """Valida o estado do modelo. Subclasses podem sobrescrever."""

    def para_dicionario(
        self,
        remover_none: bool = False,
        remover_vazios: bool = False,
        mascarar_sensiveis: bool = False,
    ) -> dict[str, Any]:
        """Converte o modelo para dicionario com filtros opcionais."""
        for nome, valor in {
            "remover_none": remover_none,
            "remover_vazios": remover_vazios,
            "mascarar_sensiveis": mascarar_sensiveis,
        }.items():
            if not isinstance(valor, bool):
                raise TypeError(f"O parametro '{nome}' deve ser booleano.")

        resultado = {
            campo.name: _serializar_valor(getattr(self, campo.name))
            for campo in fields(self)
        }

        if remover_none:
            resultado = {chave: valor for chave, valor in resultado.items() if valor is not None}

        if remover_vazios:
            resultado = {
                chave: valor
                for chave, valor in resultado.items()
                if valor not in ("", [], {}, ())
            }

        if mascarar_sensiveis:
            for campo in self.CAMPOS_SENSIVEIS:
                if campo in resultado and resultado[campo]:
                    resultado[campo] = _mascarar_texto(str(resultado[campo]))

        return resultado

    def copiar(self: TModelo, **alteracoes: Any) -> TModelo:
        """Cria uma copia do modelo aplicando alteracoes informadas."""
        return replace(self, **alteracoes)

    @classmethod
    def de_dicionario(cls: type[TModelo], dados: Mapping[str, Any]) -> TModelo:
        """Cria uma instancia utilizando apenas campos conhecidos."""
        if not isinstance(dados, Mapping):
            raise TypeError("O parametro 'dados' deve ser um mapeamento.")
        nomes = {campo.name for campo in fields(cls)}
        valores = {chave: valor for chave, valor in dados.items() if chave in nomes}
        return cls(**valores)


def _validar_texto(
    valor: str,
    nome: str,
    permitir_vazio: bool = True,
) -> str:
    """Valida uma propriedade textual e remove espacos externos."""
    if not isinstance(valor, str):
        raise TypeError(f"O campo '{nome}' deve ser uma string.")
    texto = valor.strip()
    if not texto and not permitir_vazio:
        raise ModeloValidacaoError(f"O campo '{nome}' nao pode estar vazio.")
    return texto


def _validar_email(email: str, permitir_vazio: bool = True) -> str:
    """Valida e normaliza um endereco de e-mail."""
    texto = _validar_texto(email, "email", permitir_vazio)
    if not texto:
        return ""
    if "@" not in texto or texto.startswith("@") or texto.endswith("@"):
        raise ModeloValidacaoError(f"Endereco de e-mail invalido: '{email}'.")
    return texto.lower()


def _validar_inteiro(valor: int, nome: str, minimo: int | None = None) -> int:
    """Valida um valor inteiro e seu limite inferior opcional."""
    if isinstance(valor, bool) or not isinstance(valor, int):
        raise TypeError(f"O campo '{nome}' deve ser inteiro.")
    if minimo is not None and valor < minimo:
        raise ModeloValidacaoError(
            f"O campo '{nome}' deve ser maior ou igual a {minimo}."
        )
    return valor


def _normalizar_lista_textos(valores: list[str], nome: str) -> list[str]:
    """Normaliza uma lista textual e remove itens duplicados."""
    if not isinstance(valores, list):
        raise TypeError(f"O campo '{nome}' deve ser uma lista.")
    resultado: list[str] = []
    for valor in valores:
        texto = _validar_texto(valor, nome, permitir_vazio=False)
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def _serializar_valor(valor: Any) -> Any:
    """Converte valores de modelos para estruturas serializaveis."""
    if isinstance(valor, ModeloBase):
        return valor.para_dicionario()
    if isinstance(valor, IntEnum):
        return int(valor)
    if isinstance(valor, Path):
        return str(valor)
    if isinstance(valor, (datetime, date)):
        return valor.isoformat()
    if isinstance(valor, list):
        return [_serializar_valor(item) for item in valor]
    if isinstance(valor, tuple):
        return tuple(_serializar_valor(item) for item in valor)
    if isinstance(valor, dict):
        return {chave: _serializar_valor(item) for chave, item in valor.items()}
    return valor


def _mascarar_texto(texto: str) -> str:
    """Mascara parte de um texto utilizado em logs."""
    if len(texto) <= 4:
        return "*" * len(texto)
    return f"{texto[:2]}{'*' * (len(texto) - 4)}{texto[-2:]}"


@dataclass(slots=True)
class AnexoModelo(ModeloBase):
    """Representa os dados de um anexo do Outlook."""

    nome: str
    caminho: Path | str | None = None
    tamanho_bytes: int = 0
    tipo: int | None = None
    indice: int | None = None
    salvo: bool = False

    def validar(self) -> None:
        """Valida nome, caminho, tamanho, indice e estado do anexo."""
        self.nome = _validar_texto(self.nome, "nome", permitir_vazio=False)
        if self.caminho is not None:
            if not isinstance(self.caminho, (str, Path)):
                raise TypeError("O campo 'caminho' deve ser string, Path ou None.")
            self.caminho = Path(self.caminho).expanduser()
        self.tamanho_bytes = _validar_inteiro(self.tamanho_bytes, "tamanho_bytes", 0)
        if self.indice is not None:
            self.indice = _validar_inteiro(self.indice, "indice", 1)
        if not isinstance(self.salvo, bool):
            raise TypeError("O campo 'salvo' deve ser booleano.")

    @property
    def extensao(self) -> str:
        """Retorna a extensao do nome em letras minusculas."""
        return Path(self.nome).suffix.lower()

    @property
    def tamanho_kb(self) -> float:
        """Retorna o tamanho do anexo em quilobytes."""
        return self.tamanho_bytes / 1024


@dataclass(slots=True)
class ContaModelo(ModeloBase):
    """Representa uma conta configurada no perfil Outlook."""

    CAMPOS_SENSIVEIS: ClassVar[set[str]] = {"email", "usuario"}

    nome: str = ""
    email: str = ""
    usuario: str = ""
    tipo: int | None = None
    store_id: str | None = None
    store_nome: str = ""
    indice: int | None = None
    padrao: bool = False

    def validar(self) -> None:
        """Valida os dados principais da conta."""
        self.nome = _validar_texto(self.nome, "nome")
        self.email = _validar_email(self.email)
        self.usuario = _validar_texto(self.usuario, "usuario")
        self.store_nome = _validar_texto(self.store_nome, "store_nome")
        if self.store_id is not None:
            self.store_id = _validar_texto(self.store_id, "store_id", False)
        if self.indice is not None:
            self.indice = _validar_inteiro(self.indice, "indice", 1)
        if not isinstance(self.padrao, bool):
            raise TypeError("O campo 'padrao' deve ser booleano.")


@dataclass(slots=True)
class DestinatarioModelo(ModeloBase):
    """Representa um destinatario de e-mail ou participante de reuniao."""

    CAMPOS_SENSIVEIS: ClassVar[set[str]] = {"endereco"}

    endereco: str
    nome: str = ""
    tipo: TipoDestinatarioModelo = TipoDestinatarioModelo.PARA
    resolvido: bool = False

    def validar(self) -> None:
        """Valida endereco, nome, tipo e estado de resolucao."""
        self.endereco = _validar_texto(self.endereco, "endereco", False)
        self.nome = _validar_texto(self.nome, "nome")
        if not isinstance(self.tipo, TipoDestinatarioModelo):
            try:
                self.tipo = TipoDestinatarioModelo(int(self.tipo))
            except (TypeError, ValueError) as error:
                raise ModeloValidacaoError("Tipo de destinatario invalido.") from error
        if not isinstance(self.resolvido, bool):
            raise TypeError("O campo 'resolvido' deve ser booleano.")


@dataclass(slots=True)
class EmailModelo(ModeloBase):
    """Representa os dados de uma mensagem de e-mail."""

    CAMPOS_SENSIVEIS: ClassVar[set[str]] = {"remetente_email"}

    entry_id: str | None = None
    assunto: str = ""
    remetente_nome: str = ""
    remetente_email: str = ""
    destinatarios: list[DestinatarioModelo] = field(default_factory=list)
    corpo: str = ""
    corpo_html: str = ""
    recebido_em: datetime | None = None
    enviado_em: datetime | None = None
    criado_em: datetime | None = None
    nao_lido: bool = False
    importancia: int = 1
    categorias: list[str] = field(default_factory=list)
    conversation_id: str | None = None
    message_class: str = "IPM.Note"
    anexos: list[AnexoModelo] = field(default_factory=list)

    def validar(self) -> None:
        """Valida propriedades textuais, listas, datas e importancia."""
        self.assunto = _validar_texto(self.assunto, "assunto")
        self.remetente_nome = _validar_texto(self.remetente_nome, "remetente_nome")
        self.remetente_email = _validar_email(self.remetente_email)
        self.corpo = _validar_texto(self.corpo, "corpo")
        self.corpo_html = _validar_texto(self.corpo_html, "corpo_html")
        self.message_class = _validar_texto(self.message_class, "message_class", False)
        self.importancia = _validar_inteiro(self.importancia, "importancia", 0)
        if self.importancia not in {0, 1, 2}:
            raise ModeloValidacaoError("A importancia deve ser 0, 1 ou 2.")
        if not isinstance(self.nao_lido, bool):
            raise TypeError("O campo 'nao_lido' deve ser booleano.")
        self.categorias = _normalizar_lista_textos(self.categorias, "categorias")
        if not isinstance(self.destinatarios, list) or not all(
            isinstance(item, DestinatarioModelo) for item in self.destinatarios
        ):
            raise TypeError("O campo 'destinatarios' deve conter DestinatarioModelo.")
        if not isinstance(self.anexos, list) or not all(
            isinstance(item, AnexoModelo) for item in self.anexos
        ):
            raise TypeError("O campo 'anexos' deve conter AnexoModelo.")
        for nome in ("recebido_em", "enviado_em", "criado_em"):
            valor = getattr(self, nome)
            if valor is not None and not isinstance(valor, datetime):
                raise TypeError(f"O campo '{nome}' deve ser datetime ou None.")

    @property
    def possui_anexos(self) -> bool:
        """Indica se o e-mail possui ao menos um anexo."""
        return bool(self.anexos)

    @property
    def quantidade_anexos(self) -> int:
        """Retorna a quantidade de anexos associados."""
        return len(self.anexos)

    @property
    def para(self) -> list[str]:
        """Retorna os enderecos do tipo Para."""
        return [
            item.endereco
            for item in self.destinatarios
            if item.tipo == TipoDestinatarioModelo.PARA
        ]

    @property
    def copias(self) -> list[str]:
        """Retorna os enderecos do tipo CC."""
        return [
            item.endereco
            for item in self.destinatarios
            if item.tipo == TipoDestinatarioModelo.CC
        ]

    @property
    def copias_ocultas(self) -> list[str]:
        """Retorna os enderecos do tipo CCO."""
        return [
            item.endereco
            for item in self.destinatarios
            if item.tipo == TipoDestinatarioModelo.CCO
        ]


@dataclass(slots=True)
class ContatoModelo(ModeloBase):
    """Representa os principais dados de um contato Outlook."""

    CAMPOS_SENSIVEIS: ClassVar[set[str]] = {"emails", "celular"}

    entry_id: str | None = None
    nome_completo: str = ""
    primeiro_nome: str = ""
    nome_meio: str = ""
    sobrenome: str = ""
    empresa: str = ""
    cargo: str = ""
    departamento: str = ""
    emails: list[str] = field(default_factory=list)
    telefone_comercial: str = ""
    celular: str = ""
    telefone_residencial: str = ""
    endereco_comercial: str = ""
    endereco_residencial: str = ""
    aniversario: date | datetime | None = None
    categorias: list[str] = field(default_factory=list)
    observacoes: str = ""
    possui_foto: bool = False

    def validar(self) -> None:
        """Valida os dados pessoais e profissionais do contato."""
        for nome in (
            "nome_completo", "primeiro_nome", "nome_meio", "sobrenome",
            "empresa", "cargo", "departamento", "telefone_comercial",
            "celular", "telefone_residencial", "endereco_comercial",
            "endereco_residencial", "observacoes",
        ):
            setattr(self, nome, _validar_texto(getattr(self, nome), nome))
        if not isinstance(self.emails, list):
            raise TypeError("O campo 'emails' deve ser uma lista.")
        self.emails = list(dict.fromkeys(_validar_email(email, False) for email in self.emails))
        self.categorias = _normalizar_lista_textos(self.categorias, "categorias")
        if self.aniversario is not None and not isinstance(
            self.aniversario, (date, datetime)
        ):
            raise TypeError("O campo 'aniversario' deve ser date, datetime ou None.")
        if not isinstance(self.possui_foto, bool):
            raise TypeError("O campo 'possui_foto' deve ser booleano.")


@dataclass(slots=True)
class CompromissoModelo(ModeloBase):
    """Representa um compromisso ou reuniao do calendario."""

    entry_id: str | None = None
    assunto: str = ""
    corpo: str = ""
    local: str = ""
    inicio: datetime | None = None
    fim: datetime | None = None
    duracao_minutos: int | None = None
    dia_inteiro: bool = False
    lembrete_ativo: bool = True
    lembrete_minutos: int = 15
    ocupacao: int = 2
    importancia: int = 1
    categorias: list[str] = field(default_factory=list)
    participantes: list[DestinatarioModelo] = field(default_factory=list)
    recorrente: bool = False

    def validar(self) -> None:
        """Valida datas, duracao, lembrete e participantes."""
        self.assunto = _validar_texto(self.assunto, "assunto")
        self.corpo = _validar_texto(self.corpo, "corpo")
        self.local = _validar_texto(self.local, "local")
        if self.inicio is not None and not isinstance(self.inicio, datetime):
            raise TypeError("O campo 'inicio' deve ser datetime ou None.")
        if self.fim is not None and not isinstance(self.fim, datetime):
            raise TypeError("O campo 'fim' deve ser datetime ou None.")
        if self.inicio is not None and self.fim is not None and self.fim <= self.inicio:
            raise ModeloValidacaoError("O fim deve ser posterior ao inicio.")
        if self.duracao_minutos is not None:
            self.duracao_minutos = _validar_inteiro(
                self.duracao_minutos, "duracao_minutos", 1
            )
        elif self.inicio is not None and self.fim is not None:
            self.duracao_minutos = int((self.fim - self.inicio).total_seconds() / 60)
        self.lembrete_minutos = _validar_inteiro(
            self.lembrete_minutos, "lembrete_minutos", 0
        )
        self.importancia = _validar_inteiro(self.importancia, "importancia", 0)
        self.ocupacao = _validar_inteiro(self.ocupacao, "ocupacao", 0)
        for nome in ("dia_inteiro", "lembrete_ativo", "recorrente"):
            if not isinstance(getattr(self, nome), bool):
                raise TypeError(f"O campo '{nome}' deve ser booleano.")
        self.categorias = _normalizar_lista_textos(self.categorias, "categorias")
        if not isinstance(self.participantes, list) or not all(
            isinstance(item, DestinatarioModelo) for item in self.participantes
        ):
            raise TypeError("O campo 'participantes' deve conter DestinatarioModelo.")


@dataclass(slots=True)
class ResultadoOperacao(ModeloBase):
    """Representa o resultado padronizado de uma operacao de automacao."""

    sucesso: bool
    mensagem: str = ""
    status: StatusOperacao = StatusOperacao.PENDENTE
    dados: dict[str, Any] = field(default_factory=dict)
    erros: list[str] = field(default_factory=list)
    inicio: datetime | None = None
    fim: datetime | None = None

    def validar(self) -> None:
        """Valida o resultado, status, dados, erros e datas."""
        if not isinstance(self.sucesso, bool):
            raise TypeError("O campo 'sucesso' deve ser booleano.")
        self.mensagem = _validar_texto(self.mensagem, "mensagem")
        if not isinstance(self.status, StatusOperacao):
            try:
                self.status = StatusOperacao(int(self.status))
            except (TypeError, ValueError) as error:
                raise ModeloValidacaoError("Status de operacao invalido.") from error
        if not isinstance(self.dados, dict):
            raise TypeError("O campo 'dados' deve ser um dicionario.")
        self.erros = _normalizar_lista_textos(self.erros, "erros")
        for nome in ("inicio", "fim"):
            valor = getattr(self, nome)
            if valor is not None and not isinstance(valor, datetime):
                raise TypeError(f"O campo '{nome}' deve ser datetime ou None.")
        if self.inicio is not None and self.fim is not None and self.fim < self.inicio:
            raise ModeloValidacaoError("O fim nao pode ser anterior ao inicio.")

    @property
    def duracao_segundos(self) -> float | None:
        """Retorna a duracao da operacao quando inicio e fim existem."""
        if self.inicio is None or self.fim is None:
            return None
        return (self.fim - self.inicio).total_seconds()

    @classmethod
    def com_sucesso(
        cls,
        mensagem: str = "Operacao concluida com sucesso.",
        dados: dict[str, Any] | None = None,
    ) -> ResultadoOperacao:
        """Cria um resultado de sucesso."""
        return cls(
            sucesso=True,
            mensagem=mensagem,
            status=StatusOperacao.SUCESSO,
            dados={} if dados is None else dados,
        )

    @classmethod
    def com_erro(
        cls,
        mensagem: str,
        erro: BaseException | str | None = None,
        dados: dict[str, Any] | None = None,
    ) -> ResultadoOperacao:
        """Cria um resultado de erro com detalhe opcional."""
        erros = [] if erro is None else [str(erro)]
        return cls(
            sucesso=False,
            mensagem=mensagem,
            status=StatusOperacao.ERRO,
            dados={} if dados is None else dados,
            erros=erros,
        )


__all__ = [
    "ModeloValidacaoError",
    "StatusOperacao",
    "TipoDestinatarioModelo",
    "ModeloBase",
    "AnexoModelo",
    "ContaModelo",
    "DestinatarioModelo",
    "EmailModelo",
    "ContatoModelo",
    "CompromissoModelo",
    "ResultadoOperacao",
]
