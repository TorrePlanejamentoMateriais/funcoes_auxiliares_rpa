"""
============================================================
Modulo: OUTLOOK / excecoes.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Centraliza as excecoes utilizadas pela biblioteca de automacao do
    Outlook classico. As excecoes permitem codigo, contexto, causa
    original e conversao para dicionario para facilitar logs de RPA.
Dependencias:
    - Nenhuma dependencia externa
Historico:
    v1.0.0 - 28/09/2026
        - Criacao da excecao base AutomacaoError.
        - Inclusao das excecoes dos modulos Outlook.
        - Inclusao de contexto, codigo e serializacao para logs.
============================================================
"""

from __future__ import annotations

from typing import Any, Mapping


class AutomacaoError(Exception):
    """Excecao principal da biblioteca de automacao.

    Args:
        mensagem: Descricao legivel do erro.
        codigo: Codigo opcional para logs e tratamentos condicionais.
        contexto: Dados adicionais relacionados a operacao.
        causa: Excecao original que provocou o erro.
    """

    codigo_padrao = "AUTOMACAO_ERROR"

    def __init__(
        self,
        mensagem: str | None = None,
        *,
        codigo: str | None = None,
        contexto: Mapping[str, Any] | None = None,
        causa: BaseException | None = None,
    ) -> None:
        self.mensagem = self._validar_mensagem(mensagem)
        self.codigo = self._validar_codigo(codigo or self.codigo_padrao)
        self.contexto = self._validar_contexto(contexto)
        self.causa = self._validar_causa(causa)
        super().__init__(self.mensagem)

    @staticmethod
    def _validar_mensagem(mensagem: str | None) -> str:
        """Valida a mensagem e fornece um texto padrao quando necessario."""
        if mensagem is None:
            return "Ocorreu um erro durante a automacao."
        if not isinstance(mensagem, str):
            raise TypeError("O parametro 'mensagem' deve ser string ou None.")
        texto = mensagem.strip()
        if not texto:
            return "Ocorreu um erro durante a automacao."
        return texto

    @staticmethod
    def _validar_codigo(codigo: str) -> str:
        """Valida e normaliza o codigo do erro."""
        if not isinstance(codigo, str):
            raise TypeError("O parametro 'codigo' deve ser uma string.")
        texto = codigo.strip().upper().replace(" ", "_")
        if not texto:
            raise ValueError("O parametro 'codigo' nao pode estar vazio.")
        return texto

    @staticmethod
    def _validar_contexto(
        contexto: Mapping[str, Any] | None,
    ) -> dict[str, Any]:
        """Valida e copia o contexto para evitar alteracoes externas."""
        if contexto is None:
            return {}
        if not isinstance(contexto, Mapping):
            raise TypeError("O parametro 'contexto' deve ser um mapeamento ou None.")
        return dict(contexto)

    @staticmethod
    def _validar_causa(causa: BaseException | None) -> BaseException | None:
        """Valida a excecao original associada ao erro."""
        if causa is not None and not isinstance(causa, BaseException):
            raise TypeError("O parametro 'causa' deve ser uma excecao ou None.")
        return causa

    def adicionar_contexto(self, **dados: Any) -> AutomacaoError:
        """Adiciona dados ao contexto e retorna a propria excecao."""
        self.contexto.update(dados)
        return self

    def para_dicionario(self, incluir_causa: bool = True) -> dict[str, Any]:
        """Converte a excecao para uma estrutura adequada para logs."""
        if not isinstance(incluir_causa, bool):
            raise TypeError("O parametro 'incluir_causa' deve ser booleano.")

        resultado: dict[str, Any] = {
            "tipo": type(self).__name__,
            "codigo": self.codigo,
            "mensagem": self.mensagem,
            "contexto": dict(self.contexto),
        }

        if incluir_causa:
            resultado["causa"] = (
                {
                    "tipo": type(self.causa).__name__,
                    "mensagem": str(self.causa),
                }
                if self.causa is not None
                else None
            )

        return resultado

    def __str__(self) -> str:
        """Retorna uma representacao curta e legivel do erro."""
        texto = f"[{self.codigo}] {self.mensagem}"
        if self.contexto:
            detalhes = ", ".join(
                f"{chave}={valor!r}" for chave, valor in self.contexto.items()
            )
            texto = f"{texto} | {detalhes}"
        return texto


class ValidacaoError(AutomacaoError):
    """Erro base para parametros ou dados de entrada invalidos."""

    codigo_padrao = "VALIDACAO_ERROR"


class ConfiguracaoError(AutomacaoError):
    """Erro base para configuracoes ausentes ou invalidas."""

    codigo_padrao = "CONFIGURACAO_ERROR"


class DependenciaError(AutomacaoError):
    """Erro base para dependencias externas indisponiveis."""

    codigo_padrao = "DEPENDENCIA_ERROR"


class ArquivoAutomacaoError(AutomacaoError):
    """Erro base para leitura, gravacao ou validacao de arquivos."""

    codigo_padrao = "ARQUIVO_ERROR"


class OutlookError(AutomacaoError):
    """Erro base de todos os modulos relacionados ao Outlook."""

    codigo_padrao = "OUTLOOK_ERROR"


class OutlookConexaoError(OutlookError):
    """Erro ao conectar, inicializar ou desconectar o Outlook."""

    codigo_padrao = "OUTLOOK_CONEXAO_ERROR"


class OutlookNaoDisponivelError(OutlookConexaoError):
    """O Outlook classico nao esta disponivel para automacao."""

    codigo_padrao = "OUTLOOK_NAO_DISPONIVEL"


class OutlookNaoAbertoError(OutlookConexaoError):
    """Nenhuma instancia ativa do Outlook foi localizada."""

    codigo_padrao = "OUTLOOK_NAO_ABERTO"


class PyWin32NaoDisponivelError(DependenciaError, OutlookError):
    """A biblioteca pywin32 nao esta instalada ou acessivel."""

    codigo_padrao = "PYWIN32_NAO_DISPONIVEL"


class NamespaceMAPIError(OutlookConexaoError):
    """O Namespace MAPI nao pode ser obtido ou inicializado."""

    codigo_padrao = "NAMESPACE_MAPI_ERROR"


class PerfilOutlookError(OutlookConexaoError):
    """O perfil do Outlook nao pode ser acessado ou autenticado."""

    codigo_padrao = "PERFIL_OUTLOOK_ERROR"


class OutlookContaError(OutlookError):
    """Erro base para operacoes relacionadas a contas do Outlook."""

    codigo_padrao = "OUTLOOK_CONTA_ERROR"


class ContaNaoEncontradaError(OutlookContaError):
    """A conta solicitada nao foi encontrada no perfil atual."""

    codigo_padrao = "CONTA_NAO_ENCONTRADA"


class ContaInvalidaError(OutlookContaError, ValidacaoError):
    """O objeto ou os dados da conta sao invalidos."""

    codigo_padrao = "CONTA_INVALIDA"


class ContaEnvioError(OutlookContaError):
    """A conta nao pode ser utilizada para enviar o item."""

    codigo_padrao = "CONTA_ENVIO_ERROR"


class OutlookPastaError(OutlookError):
    """Erro base para operacoes relacionadas a pastas do Outlook."""

    codigo_padrao = "OUTLOOK_PASTA_ERROR"


class PastaNaoEncontradaError(OutlookPastaError):
    """A pasta solicitada nao foi encontrada."""

    codigo_padrao = "PASTA_NAO_ENCONTRADA"


class OutlookEmailError(OutlookError):
    """Erro base para operacoes relacionadas a mensagens de e-mail."""

    codigo_padrao = "OUTLOOK_EMAIL_ERROR"


class EmailCriacaoError(OutlookEmailError):
    """O Outlook nao conseguiu criar o MailItem."""

    codigo_padrao = "EMAIL_CRIACAO_ERROR"


class EmailEnvioError(OutlookEmailError):
    """O Outlook nao conseguiu enviar o MailItem."""

    codigo_padrao = "EMAIL_ENVIO_ERROR"


class EmailNaoEncontradoError(OutlookEmailError):
    """Nenhuma mensagem correspondente foi encontrada."""

    codigo_padrao = "EMAIL_NAO_ENCONTRADO"


class EmailSalvamentoError(OutlookEmailError):
    """O Outlook nao conseguiu salvar o MailItem."""

    codigo_padrao = "EMAIL_SALVAMENTO_ERROR"


class DestinatarioInvalidoError(OutlookEmailError, ValidacaoError):
    """O destinatario informado e invalido ou nao foi resolvido."""

    codigo_padrao = "DESTINATARIO_INVALIDO"


class OutlookAnexoError(OutlookError):
    """Erro base para operacoes relacionadas a anexos."""

    codigo_padrao = "OUTLOOK_ANEXO_ERROR"


class AnexoNaoEncontradoError(OutlookAnexoError):
    """O anexo solicitado nao foi encontrado."""

    codigo_padrao = "ANEXO_NAO_ENCONTRADO"


class AnexoInvalidoError(OutlookAnexoError, ValidacaoError):
    """O anexo nao atende aos criterios de validacao."""

    codigo_padrao = "ANEXO_INVALIDO"


class AnexoNaoSalvoError(OutlookAnexoError):
    """O anexo nao pode ser salvo no sistema de arquivos."""

    codigo_padrao = "ANEXO_NAO_SALVO"


class OutlookCalendarioError(OutlookError):
    """Erro base para compromissos e reunioes do Outlook."""

    codigo_padrao = "OUTLOOK_CALENDARIO_ERROR"


class CompromissoNaoEncontradoError(OutlookCalendarioError):
    """O compromisso solicitado nao foi encontrado."""

    codigo_padrao = "COMPROMISSO_NAO_ENCONTRADO"


class CompromissoEnvioError(OutlookCalendarioError):
    """O convite ou cancelamento nao pode ser enviado."""

    codigo_padrao = "COMPROMISSO_ENVIO_ERROR"


class RecorrenciaError(OutlookCalendarioError):
    """O padrao de recorrencia nao pode ser configurado."""

    codigo_padrao = "RECORRENCIA_ERROR"


class OutlookContatoError(OutlookError):
    """Erro base para operacoes relacionadas a contatos."""

    codigo_padrao = "OUTLOOK_CONTATO_ERROR"


class ContatoNaoEncontradoError(OutlookContatoError):
    """O contato solicitado nao foi encontrado."""

    codigo_padrao = "CONTATO_NAO_ENCONTRADO"


class ContatoInvalidoError(OutlookContatoError, ValidacaoError):
    """O contato ou seus dados sao invalidos."""

    codigo_padrao = "CONTATO_INVALIDO"


class ContatoSalvamentoError(OutlookContatoError):
    """O contato nao pode ser salvo pelo Outlook."""

    codigo_padrao = "CONTATO_SALVAMENTO_ERROR"


class OutlookPesquisaError(OutlookError):
    """Erro ao pesquisar ou filtrar itens do Outlook."""

    codigo_padrao = "OUTLOOK_PESQUISA_ERROR"


class OutlookRespostaError(OutlookError):
    """Erro ao responder ou encaminhar uma mensagem."""

    codigo_padrao = "OUTLOOK_RESPOSTA_ERROR"


class OutlookRegraError(OutlookError):
    """Erro ao aplicar regras, categorias ou marcacoes."""

    codigo_padrao = "OUTLOOK_REGRA_ERROR"


__all__ = [
    "AutomacaoError",
    "ValidacaoError",
    "ConfiguracaoError",
    "DependenciaError",
    "ArquivoAutomacaoError",
    "OutlookError",
    "OutlookConexaoError",
    "OutlookNaoDisponivelError",
    "OutlookNaoAbertoError",
    "PyWin32NaoDisponivelError",
    "NamespaceMAPIError",
    "PerfilOutlookError",
    "OutlookContaError",
    "ContaNaoEncontradaError",
    "ContaInvalidaError",
    "ContaEnvioError",
    "OutlookPastaError",
    "PastaNaoEncontradaError",
    "OutlookEmailError",
    "EmailCriacaoError",
    "EmailEnvioError",
    "EmailNaoEncontradoError",
    "EmailSalvamentoError",
    "DestinatarioInvalidoError",
    "OutlookAnexoError",
    "AnexoNaoEncontradoError",
    "AnexoInvalidoError",
    "AnexoNaoSalvoError",
    "OutlookCalendarioError",
    "CompromissoNaoEncontradoError",
    "CompromissoEnvioError",
    "RecorrenciaError",
    "OutlookContatoError",
    "ContatoNaoEncontradoError",
    "ContatoInvalidoError",
    "ContatoSalvamentoError",
    "OutlookPesquisaError",
    "OutlookRespostaError",
    "OutlookRegraError",
]
