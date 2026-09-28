"""
============================================================
Modulo: OUTLOOK / contatos.py
Autor: Tiago Eneas Antunes (Z701038)
Data de Criacao: 28/09/2026
Ultima Alteracao: 28/09/2026
Versao: 1.0.0
Descricao:
    Funcoes auxiliares para criar, configurar, consultar, pesquisar,
    salvar, atualizar, mover e excluir contatos no Outlook classico.
Dependencias:
    - Outlook classico para Windows
    - constantes.py
Historico:
    v1.0.0 - 28/09/2026
        - Criacao inicial do modulo.
        - Inclusao de criacao e configuracao de contatos.
        - Inclusao de pesquisa por nome, e-mail e empresa.
        - Inclusao de telefones, enderecos, foto e categorias.
============================================================
"""

from datetime import date, datetime
from pathlib import Path
from typing import Any

try:
    from excecoes import AutomacaoError
except ImportError:
    class AutomacaoError(Exception):
        """Excecao base usada quando o modulo central nao esta disponivel."""

try:
    from constantes import OL_CONTACT_ITEM, OL_FOLDER_CONTACTS
except ImportError:
    OL_CONTACT_ITEM = 2
    OL_FOLDER_CONTACTS = 10


class OutlookContatosError(AutomacaoError):
    """Erro base para operacoes relacionadas a contatos do Outlook."""


class ContatoInvalidoError(OutlookContatosError):
    """O ContactItem informado esta ausente ou possui dados invalidos."""


class ContatoNaoEncontradoError(OutlookContatosError):
    """Nenhum contato correspondente foi encontrado."""


class ContatoCriacaoError(OutlookContatosError):
    """O contato nao pode ser criado pelo Outlook."""


class ContatoSalvamentoError(OutlookContatosError):
    """O contato nao pode ser salvo pelo Outlook."""


class ContatoImagemError(OutlookContatosError):
    """A imagem do contato nao pode ser adicionada ou removida."""


class PastaContatosError(OutlookContatosError):
    """A pasta de contatos nao pode ser obtida ou consultada."""


def _obter_propriedade(objeto: Any, nome: str, padrao: Any = None) -> Any:
    """Obtem uma propriedade COM, retornando um valor padrao em erro."""
    try:
        return getattr(objeto, nome)
    except Exception:
        return padrao


def _validar_objeto(objeto: Any, nome: str) -> None:
    """Valida se um objeto obrigatorio foi informado."""
    if objeto is None:
        raise ContatoInvalidoError(f"O parametro '{nome}' nao pode ser None.")


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


def _normalizar_email(email: str, permitir_vazio: bool = False) -> str:
    """Valida e normaliza um endereco de e-mail."""
    texto = _validar_texto(email, "email", permitir_vazio=permitir_vazio)
    if not texto:
        return ""
    if "@" not in texto or texto.startswith("@") or texto.endswith("@"):
        raise ValueError("O endereco de e-mail informado e invalido.")
    return texto.lower()


def _normalizar_lista_textos(
    valores: str | list[str] | tuple[str, ...] | set[str],
    nome: str,
) -> list[str]:
    """Normaliza um texto ou uma colecao de textos sem duplicidades."""
    if isinstance(valores, str):
        valores = [valores]
    if not isinstance(valores, (list, tuple, set)) or not valores:
        raise TypeError(f"O parametro '{nome}' deve ser uma colecao nao vazia.")
    resultado = []
    for valor in valores:
        texto = _validar_texto(valor, nome)
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def _converter_data(valor: date | datetime | None, nome: str) -> date | datetime | None:
    """Valida valores de data opcionais aceitos por ContactItem."""
    if valor is None or isinstance(valor, (date, datetime)):
        return valor
    raise TypeError(f"O parametro '{nome}' deve ser date, datetime ou None.")


def _definir_propriedade(contato: Any, propriedade: str, valor: Any) -> None:
    """Define uma propriedade de ContactItem com tratamento padronizado."""
    _validar_objeto(contato, "contato")
    try:
        setattr(contato, propriedade, valor)
    except Exception as error:
        raise ContatoInvalidoError(
            f"Nao foi possivel definir a propriedade '{propriedade}'."
        ) from error


def _comparar_texto(
    atual: str,
    esperado: str,
    correspondencia_exata: bool,
    considerar_maiusculas: bool,
) -> bool:
    """Compara textos de forma exata ou parcial."""
    if not isinstance(correspondencia_exata, bool):
        raise TypeError("O parametro 'correspondencia_exata' deve ser booleano.")
    if not isinstance(considerar_maiusculas, bool):
        raise TypeError("O parametro 'considerar_maiusculas' deve ser booleano.")
    a = atual if considerar_maiusculas else atual.lower()
    e = esperado if considerar_maiusculas else esperado.lower()
    return a == e if correspondencia_exata else e in a


def obter_pasta_contatos(namespace: Any) -> Any:
    """Retorna a pasta padrao de contatos do Namespace MAPI."""
    _validar_objeto(namespace, "namespace")
    try:
        return namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    except Exception as error:
        raise PastaContatosError(
            "Nao foi possivel obter a pasta padrao de contatos."
        ) from error


def criar_contato(outlook: Any) -> Any:
    """Cria e retorna um novo ContactItem ainda nao salvo."""
    _validar_objeto(outlook, "outlook")
    try:
        return outlook.CreateItem(OL_CONTACT_ITEM)
    except Exception as error:
        raise ContatoCriacaoError("Nao foi possivel criar o contato.") from error


def definir_nome(
    contato: Any,
    primeiro_nome: str = "",
    nome_meio: str = "",
    sobrenome: str = "",
    nome_completo: str | None = None,
    sufixo: str = "",
) -> Any:
    """Define os componentes do nome de um contato."""
    if nome_completo is not None:
        _definir_propriedade(
            contato, "FullName", _validar_texto(nome_completo, "nome_completo")
        )
    campos = {
        "FirstName": _validar_texto(primeiro_nome, "primeiro_nome", True),
        "MiddleName": _validar_texto(nome_meio, "nome_meio", True),
        "LastName": _validar_texto(sobrenome, "sobrenome", True),
        "Suffix": _validar_texto(sufixo, "sufixo", True),
    }
    if nome_completo is None and not any(campos.values()):
        raise ValueError("Informe nome_completo ou ao menos uma parte do nome.")
    for propriedade, valor in campos.items():
        _definir_propriedade(contato, propriedade, valor)
    return contato


def definir_empresa(
    contato: Any,
    empresa: str = "",
    cargo: str = "",
    departamento: str = "",
    gerente: str = "",
) -> Any:
    """Define os dados profissionais de um contato."""
    dados = {
        "CompanyName": (empresa, "empresa"),
        "JobTitle": (cargo, "cargo"),
        "Department": (departamento, "departamento"),
        "ManagerName": (gerente, "gerente"),
    }
    for propriedade, (valor, nome) in dados.items():
        _definir_propriedade(contato, propriedade, _validar_texto(valor, nome, True))
    return contato


def definir_email(
    contato: Any,
    email: str,
    posicao: int = 1,
    nome_exibicao: str | None = None,
) -> Any:
    """Define um dos tres enderecos de e-mail do contato."""
    if isinstance(posicao, bool) or not isinstance(posicao, int):
        raise TypeError("O parametro 'posicao' deve ser inteiro.")
    if posicao not in {1, 2, 3}:
        raise ValueError("O parametro 'posicao' deve ser 1, 2 ou 3.")
    endereco = _normalizar_email(email)
    _definir_propriedade(contato, f"Email{posicao}Address", endereco)
    _definir_propriedade(contato, f"Email{posicao}AddressType", "SMTP")
    if nome_exibicao is not None:
        _definir_propriedade(
            contato,
            f"Email{posicao}DisplayName",
            _validar_texto(nome_exibicao, "nome_exibicao"),
        )
    return contato


def obter_emails_contato(contato: Any, remover_vazios: bool = True) -> list[str]:
    """Retorna os tres enderecos de e-mail cadastrados no contato."""
    _validar_objeto(contato, "contato")
    if not isinstance(remover_vazios, bool):
        raise TypeError("O parametro 'remover_vazios' deve ser booleano.")
    emails = [
        str(_obter_propriedade(contato, f"Email{i}Address", "") or "").strip()
        for i in range(1, 4)
    ]
    return [email for email in emails if email] if remover_vazios else emails


def definir_telefones(
    contato: Any,
    comercial: str = "",
    comercial_secundario: str = "",
    celular: str = "",
    residencial: str = "",
    residencial_secundario: str = "",
    fax_comercial: str = "",
    fax_residencial: str = "",
) -> Any:
    """Define os principais telefones de um contato."""
    dados = {
        "BusinessTelephoneNumber": comercial,
        "Business2TelephoneNumber": comercial_secundario,
        "MobileTelephoneNumber": celular,
        "HomeTelephoneNumber": residencial,
        "Home2TelephoneNumber": residencial_secundario,
        "BusinessFaxNumber": fax_comercial,
        "HomeFaxNumber": fax_residencial,
    }
    for propriedade, valor in dados.items():
        _definir_propriedade(
            contato, propriedade, _validar_texto(valor, propriedade, True)
        )
    return contato


def definir_endereco_comercial(
    contato: Any,
    rua: str = "",
    cidade: str = "",
    estado: str = "",
    cep: str = "",
    pais: str = "",
    caixa_postal: str = "",
) -> Any:
    """Define o endereco comercial de um contato."""
    dados = {
        "BusinessAddressStreet": rua,
        "BusinessAddressCity": cidade,
        "BusinessAddressState": estado,
        "BusinessAddressPostalCode": cep,
        "BusinessAddressCountry": pais,
        "BusinessAddressPostOfficeBox": caixa_postal,
    }
    for propriedade, valor in dados.items():
        _definir_propriedade(contato, propriedade, _validar_texto(valor, propriedade, True))
    return contato


def definir_endereco_residencial(
    contato: Any,
    rua: str = "",
    cidade: str = "",
    estado: str = "",
    cep: str = "",
    pais: str = "",
    caixa_postal: str = "",
) -> Any:
    """Define o endereco residencial de um contato."""
    dados = {
        "HomeAddressStreet": rua,
        "HomeAddressCity": cidade,
        "HomeAddressState": estado,
        "HomeAddressPostalCode": cep,
        "HomeAddressCountry": pais,
        "HomeAddressPostOfficeBox": caixa_postal,
    }
    for propriedade, valor in dados.items():
        _definir_propriedade(contato, propriedade, _validar_texto(valor, propriedade, True))
    return contato


def definir_datas(
    contato: Any,
    aniversario: date | datetime | None = None,
    aniversario_casamento: date | datetime | None = None,
) -> Any:
    """Define aniversario e aniversario de casamento do contato."""
    aniversario = _converter_data(aniversario, "aniversario")
    aniversario_casamento = _converter_data(
        aniversario_casamento, "aniversario_casamento"
    )
    if aniversario is not None:
        _definir_propriedade(contato, "Birthday", aniversario)
    if aniversario_casamento is not None:
        _definir_propriedade(contato, "Anniversary", aniversario_casamento)
    return contato


def definir_categorias(
    contato: Any,
    categorias: str | list[str] | tuple[str, ...] | set[str],
) -> Any:
    """Define categorias do contato, removendo valores duplicados."""
    valores = _normalizar_lista_textos(categorias, "categorias")
    _definir_propriedade(contato, "Categories", ", ".join(valores))
    return contato


def definir_observacoes(contato: Any, observacoes: str) -> Any:
    """Define o corpo de observacoes do contato."""
    _definir_propriedade(
        contato, "Body", _validar_texto(observacoes, "observacoes", True)
    )
    return contato


def adicionar_foto(contato: Any, caminho_imagem: str | Path) -> Any:
    """Adiciona uma imagem existente ao ContactItem."""
    _validar_objeto(contato, "contato")
    if not isinstance(caminho_imagem, (str, Path)):
        raise TypeError("O parametro 'caminho_imagem' deve ser string ou Path.")
    caminho = Path(caminho_imagem).expanduser().resolve()
    if not caminho.is_file():
        raise FileNotFoundError(f"A imagem nao existe: {caminho}")
    try:
        contato.AddPicture(str(caminho))
        return contato
    except Exception as error:
        raise ContatoImagemError("Nao foi possivel adicionar a foto.") from error


def remover_foto(contato: Any) -> Any:
    """Remove a foto atual do contato."""
    _validar_objeto(contato, "contato")
    try:
        contato.RemovePicture()
        return contato
    except Exception as error:
        raise ContatoImagemError("Nao foi possivel remover a foto.") from error


def configurar_contato(
    contato: Any,
    nome_completo: str,
    email: str | None = None,
    empresa: str = "",
    cargo: str = "",
    celular: str = "",
    telefone_comercial: str = "",
    categorias: str | list[str] | tuple[str, ...] | set[str] | None = None,
    observacoes: str = "",
) -> Any:
    """Configura as principais propriedades de um contato corporativo."""
    definir_nome(contato, nome_completo=nome_completo)
    definir_empresa(contato, empresa=empresa, cargo=cargo)
    definir_telefones(contato, comercial=telefone_comercial, celular=celular)
    definir_observacoes(contato, observacoes)
    if email is not None:
        definir_email(contato, email, posicao=1)
    if categorias is not None:
        definir_categorias(contato, categorias)
    return contato


def criar_contato_completo(
    outlook: Any,
    nome_completo: str,
    salvar: bool = True,
    exibir: bool = False,
    **configuracoes: Any,
) -> Any:
    """Cria, configura e opcionalmente salva ou exibe um contato."""
    if not isinstance(salvar, bool) or not isinstance(exibir, bool):
        raise TypeError("Os parametros 'salvar' e 'exibir' devem ser booleanos.")
    contato = criar_contato(outlook)
    configurar_contato(contato, nome_completo, **configuracoes)
    if salvar:
        salvar_contato(contato)
    if exibir:
        exibir_contato(contato)
    return contato


def salvar_contato(contato: Any) -> Any:
    """Salva o ContactItem e retorna o proprio objeto."""
    _validar_objeto(contato, "contato")
    try:
        contato.Save()
        return contato
    except Exception as error:
        raise ContatoSalvamentoError("Nao foi possivel salvar o contato.") from error


def exibir_contato(contato: Any, modal: bool = False) -> Any:
    """Exibe a janela do contato no Outlook."""
    _validar_objeto(contato, "contato")
    if not isinstance(modal, bool):
        raise TypeError("O parametro 'modal' deve ser booleano.")
    try:
        contato.Display(modal)
        return contato
    except Exception as error:
        raise OutlookContatosError("Nao foi possivel exibir o contato.") from error


def excluir_contato(contato: Any) -> bool:
    """Exclui o contato da pasta onde ele esta armazenado."""
    _validar_objeto(contato, "contato")
    try:
        contato.Delete()
        return True
    except Exception as error:
        raise OutlookContatosError("Nao foi possivel excluir o contato.") from error


def mover_contato(contato: Any, pasta_destino: Any) -> Any:
    """Move um contato para outra pasta e retorna o item movido."""
    _validar_objeto(contato, "contato")
    _validar_objeto(pasta_destino, "pasta_destino")
    try:
        return contato.Move(pasta_destino)
    except Exception as error:
        raise OutlookContatosError("Nao foi possivel mover o contato.") from error


def listar_contatos(
    pasta_contatos: Any,
    ordenar_por: str = "[FullName]",
    ordem_decrescente: bool = False,
    limite: int | None = None,
) -> list[Any]:
    """Lista contatos de uma pasta, com ordenacao e limite opcionais."""
    _validar_objeto(pasta_contatos, "pasta_contatos")
    campo = _validar_texto(ordenar_por, "ordenar_por")
    if not isinstance(ordem_decrescente, bool):
        raise TypeError("O parametro 'ordem_decrescente' deve ser booleano.")
    if limite is not None:
        if isinstance(limite, bool) or not isinstance(limite, int):
            raise TypeError("O parametro 'limite' deve ser inteiro ou None.")
        if limite < 1:
            raise ValueError("O parametro 'limite' deve ser maior que zero.")
    try:
        itens = pasta_contatos.Items
        itens.Sort(campo, ordem_decrescente)
        quantidade = int(itens.Count)
    except Exception as error:
        raise PastaContatosError("Nao foi possivel listar os contatos.") from error
    quantidade = quantidade if limite is None else min(quantidade, limite)
    resultado = []
    for indice in range(1, quantidade + 1):
        try:
            resultado.append(itens.Item(indice))
        except Exception as error:
            raise PastaContatosError(
                f"Nao foi possivel obter o contato no indice {indice}."
            ) from error
    return resultado


def buscar_contatos(
    pasta_contatos: Any,
    nome: str | None = None,
    email: str | None = None,
    empresa: str | None = None,
    cargo: str | None = None,
    correspondencia_exata: bool = False,
    considerar_maiusculas: bool = False,
    limite: int | None = None,
) -> list[Any]:
    """Pesquisa contatos por nome, e-mail, empresa e cargo."""
    filtros = {
        "nome": _validar_texto(nome, "nome") if nome is not None else None,
        "email": _normalizar_email(email) if email is not None else None,
        "empresa": _validar_texto(empresa, "empresa") if empresa is not None else None,
        "cargo": _validar_texto(cargo, "cargo") if cargo is not None else None,
    }
    if not any(filtros.values()):
        raise ValueError("Informe ao menos um filtro de pesquisa.")
    if limite is not None and (isinstance(limite, bool) or not isinstance(limite, int)):
        raise TypeError("O parametro 'limite' deve ser inteiro ou None.")
    if limite is not None and limite < 1:
        raise ValueError("O parametro 'limite' deve ser maior que zero.")
    resultado = []
    for contato in listar_contatos(pasta_contatos):
        nome_atual = str(_obter_propriedade(contato, "FullName", "") or "")
        empresa_atual = str(_obter_propriedade(contato, "CompanyName", "") or "")
        cargo_atual = str(_obter_propriedade(contato, "JobTitle", "") or "")
        emails = obter_emails_contato(contato)
        if filtros["nome"] and not _comparar_texto(nome_atual, filtros["nome"], correspondencia_exata, considerar_maiusculas):
            continue
        if filtros["empresa"] and not _comparar_texto(empresa_atual, filtros["empresa"], correspondencia_exata, considerar_maiusculas):
            continue
        if filtros["cargo"] and not _comparar_texto(cargo_atual, filtros["cargo"], correspondencia_exata, considerar_maiusculas):
            continue
        if filtros["email"] and not any(
            _comparar_texto(e, filtros["email"], correspondencia_exata, considerar_maiusculas)
            for e in emails
        ):
            continue
        resultado.append(contato)
        if limite is not None and len(resultado) >= limite:
            break
    return resultado


def obter_contato_por_email(pasta_contatos: Any, email: str) -> Any:
    """Retorna o primeiro contato com o endereco de e-mail informado."""
    resultados = buscar_contatos(
        pasta_contatos,
        email=email,
        correspondencia_exata=True,
        considerar_maiusculas=False,
        limite=1,
    )
    if not resultados:
        raise ContatoNaoEncontradoError(
            f"Nenhum contato foi encontrado para o e-mail '{email}'."
        )
    return resultados[0]


def obter_contato_por_nome(
    pasta_contatos: Any,
    nome: str,
    correspondencia_exata: bool = True,
) -> Any:
    """Retorna o primeiro contato correspondente ao nome informado."""
    resultados = buscar_contatos(
        pasta_contatos,
        nome=nome,
        correspondencia_exata=correspondencia_exata,
        limite=1,
    )
    if not resultados:
        raise ContatoNaoEncontradoError(
            f"Nenhum contato correspondente a '{nome}' foi encontrado."
        )
    return resultados[0]


def contato_existe(
    pasta_contatos: Any,
    identificador: str,
    buscar_por_email: bool = True,
) -> bool:
    """Verifica se um contato existe por e-mail ou nome."""
    if not isinstance(buscar_por_email, bool):
        raise TypeError("O parametro 'buscar_por_email' deve ser booleano.")
    try:
        if buscar_por_email:
            obter_contato_por_email(pasta_contatos, identificador)
        else:
            obter_contato_por_nome(pasta_contatos, identificador)
        return True
    except ContatoNaoEncontradoError:
        return False


def obter_contato_por_entry_id(
    namespace: Any,
    entry_id: str,
    store_id: str | None = None,
) -> Any:
    """Localiza um contato por EntryID e StoreID opcional."""
    _validar_objeto(namespace, "namespace")
    identificador = _validar_texto(entry_id, "entry_id")
    if store_id is not None:
        store_id = _validar_texto(store_id, "store_id")
    try:
        contato = (
            namespace.GetItemFromID(identificador, store_id)
            if store_id is not None
            else namespace.GetItemFromID(identificador)
        )
    except Exception as error:
        raise ContatoNaoEncontradoError(
            f"Nao foi possivel localizar o contato '{identificador}'."
        ) from error
    if contato is None:
        raise ContatoNaoEncontradoError(
            f"O contato '{identificador}' nao foi encontrado."
        )
    return contato


def obter_dados_contato(contato: Any) -> dict[str, Any]:
    """Converte as principais propriedades do contato para dicionario."""
    _validar_objeto(contato, "contato")
    return {
        "entry_id": _obter_propriedade(contato, "EntryID"),
        "nome_completo": str(_obter_propriedade(contato, "FullName", "") or ""),
        "primeiro_nome": str(_obter_propriedade(contato, "FirstName", "") or ""),
        "nome_meio": str(_obter_propriedade(contato, "MiddleName", "") or ""),
        "sobrenome": str(_obter_propriedade(contato, "LastName", "") or ""),
        "empresa": str(_obter_propriedade(contato, "CompanyName", "") or ""),
        "cargo": str(_obter_propriedade(contato, "JobTitle", "") or ""),
        "departamento": str(_obter_propriedade(contato, "Department", "") or ""),
        "emails": obter_emails_contato(contato),
        "telefone_comercial": str(_obter_propriedade(contato, "BusinessTelephoneNumber", "") or ""),
        "celular": str(_obter_propriedade(contato, "MobileTelephoneNumber", "") or ""),
        "telefone_residencial": str(_obter_propriedade(contato, "HomeTelephoneNumber", "") or ""),
        "endereco_comercial": str(_obter_propriedade(contato, "BusinessAddress", "") or ""),
        "endereco_residencial": str(_obter_propriedade(contato, "HomeAddress", "") or ""),
        "aniversario": _obter_propriedade(contato, "Birthday"),
        "categorias": str(_obter_propriedade(contato, "Categories", "") or ""),
        "observacoes": str(_obter_propriedade(contato, "Body", "") or ""),
        "possui_foto": bool(_obter_propriedade(contato, "HasPicture", False)),
    }


def encaminhar_como_vcard(contato: Any, exibir: bool = False) -> Any:
    """Cria um MailItem com o contato anexado em formato vCard."""
    _validar_objeto(contato, "contato")
    if not isinstance(exibir, bool):
        raise TypeError("O parametro 'exibir' deve ser booleano.")
    try:
        email = contato.ForwardAsVcard()
        if exibir:
            email.Display(False)
        return email
    except Exception as error:
        raise OutlookContatosError(
            "Nao foi possivel encaminhar o contato como vCard."
        ) from error
