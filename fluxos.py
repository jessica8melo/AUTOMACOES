#!/usr/bin/env python3
"""
Definição dos fluxos (checklists) de conferência de documentos.

Cada fluxo representa um tipo de processo (ex.: FQ415-031 - OC Padrão com
Contrato). Dentro de cada fluxo, cada "documento" é um tipo de arquivo que
deve ser conferido (Contrato, Nota Técnica, FQ415-075, etc.), e para cada
documento há uma lista de campos que devem ser localizados e conferidos.

Esse módulo é só dados: quem decide QUAL fluxo está sendo aplicado e QUAL
documento cada arquivo representa é o main.py (com ajuda de doc_types.py).
Só depois dessas duas decisões é que os campos daqui são usados para buscar
informação dentro do arquivo.

Convenção especial:
    Alguns itens do checklist não são "campos" a extrair, e sim conferências
    de assinatura (ex.: "Conferir assinaturas de todos os documentos"). Esses
    itens são marcados com o prefixo "ASSINATURAS:" e são tratados à parte
    (usam o verificador de assinatura de pdfs.py em vez de busca de campo).
"""

ASSINATURAS_PREFIXO = "ASSINATURAS:"

FLUXOS = {
    "FQ412-033": {
        "nome": "Liberação de Acordo em Aberto",
        "documentos": {
            "Solicitação de Entrega": [
                "Valor total da solicitação",   # Total
                "Endereço",                     # Entregar Para
                "Quantidade",                   # Quantidade
                "Data do fornecimento",         # Prometido
                "Código BBTS",                  # Código
                "Especificação do Bem",         # Nome

                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
            "FQ415-075": [                      # Conta de Débito
                "UOR",
                "Conta Contábil",
            ],
            "Nota Técnica": [
                "Conta de Débito",

                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
            "FQ412-034": [
                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
            "FQ412-035": [
                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
        },
    },
    "FQ415-031": {
        "nome": "OC Padrão – Com Contrato",
        "documentos": {
            "Contrato / Aditivo": [
                "Contratada",   # Fornecedor
                "Sede na cidade de",    #Entregar Para
                "DGCO",
                "Modalidade de Contratação", #Modalidade de Contratação
                "Tipo",
                "UDM",
                "Preço",
                "Prometido",
                "Pagamento",
            ],
            "Projeto Básico": [
                "Pagamento",
                "Entregar Para",
                "Conta de Débito",
            ],
            "Nota Técnica": [
                "Número da Nota Técnica",
                "Conta de Débito",
            ],
            "FQ415-075": [
                "Fornecedor",
                "Local",
                "Descrição",
                "Entregar Para",
            ],
        },
    },
    "FQ415-055": {
        "nome": "Acordo de Compra em Aberto",
        "documentos": {
            "Contrato": [
                "Fornecedor",
                "Local",
                "Descrição",
                "Total",
                "DGCO Fornecedor SICON_GESCON",
                "Modalidade de Contratação",
                "Tipo",
                "UDM",
                "Quantidade Acordada",
                "Quantia Acordada",
                "Preço",
                "Pagamento",
            ],
            "Nota Técnica": [
                "Número da Nota Técnica",
            ],
            "Projeto Básico": [
                "Pagamento",
            ],
            "FQ415-075": [
                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
        },
    },
    "FQ415-084": {
        "nome": "Acordo de Compra do Contrato",
        "documentos": {
            "Contrato": [
                "Fornecedor",
                "Local",
                "Descrição",
                "Total",
                "Modalidade de Contratação",
                "Pagamento",
                "Frete",
                "FOB",
            ],
            "Nota Técnica": [
                "Número da Nota Técnica",
            ],
            "FQ415-075": [
                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
        },
    },
    "FQ412-043": {
        "nome": "Liberação de Acordo do Contrato",
        "documentos": {
            "ACC Master": [
                "Local",
                "Excede",
                "Número da OC Master",
            ],
            "Contrato": [
                "Fornecedor",
                "Descrição",
                "Pagamento",
                "Tipo",
            ],
            "FQ415-075": [
                "Faturar Para",
                "Entrega Para",
                "Natureza da Transação",
            ],
            "Nota Técnica": [
                "Número da Nota Técnica",
                "Conta de Débito",
            ],
            "Solicitação de Entrega": [
                "Entrega Para",
            ],
        },
    },
    "FQ412-069": {
        "nome": "Condomínio",
        "documentos": {
            "Contrato / Aditivo": [
                "Modalidade de Contratação",
                "Pagamento",
            ],
            "Ata do Condomínio": [
                f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos",
            ],
            "Boleto do Condomínio": [
                "Fornecedor",
                "Local",
                "Valor",
            ],
            "Nota Técnica": [
                "Número da Nota Técnica",
                "Conta de Débito",
            ],
            "Documento de Referência da Área": [
                "Org",
                "Entregar Para",
            ],
        },
    },
}


def listar_fluxos() -> list:
    """Devolve [(id_fluxo, nome_fluxo), ...] na ordem de definição."""
    return [(fid, dados["nome"]) for fid, dados in FLUXOS.items()]


def documentos_do_fluxo(id_fluxo: str) -> dict:
    """Devolve o dict {documento: [campos]} de um fluxo, ou {} se não existir."""
    fluxo = FLUXOS.get(id_fluxo)
    return fluxo["documentos"] if fluxo else {}


def campos_do_documento(id_fluxo: str, tipo_documento: str) -> list:
    """Devolve a lista de campos/itens de checklist de um documento dentro de
    um fluxo específico, ou [] se a combinação não existir."""
    return documentos_do_fluxo(id_fluxo).get(tipo_documento, [])


def eh_item_assinatura(campo: str) -> bool:
    """True se o item do checklist for uma conferência de assinatura, e não
    um campo a extrair do documento."""
    return campo.startswith(ASSINATURAS_PREFIXO)


def descricao_item_assinatura(campo: str) -> str:
    """Remove o prefixo especial e devolve só a descrição legível do item."""
    return campo[len(ASSINATURAS_PREFIXO):].strip()