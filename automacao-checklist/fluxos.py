#!/usr/bin/env python3
"""
Definição dos fluxos (checklists) de conferência de documentos.

Cada fluxo representa um tipo de processo (ex.: FQ415-031 - OC Padrão com
Contrato). Dentro de cada fluxo, listamos apenas os NOMES dos documentos
que fazem parte do checklist — os campos a conferir de cada documento
vêm de documentos.py, que é o ponto único de configuração da etapa de extração.

Se um documento do fluxo não existir em documentos.py (ex.: "FQ412-034",
"Ata do Condomínio"), significa que ele não tem campos a extrair — nesse
caso o único item do checklist passa a ser, automaticamente, a conferência
de assinaturas.

Convenção especial:
    Itens de checklist que são conferências de assinatura (e não campos a
    extrair) usam o prefixo "ASSINATURAS:". Esses itens são tratados à parte
    (usam o verificador de assinatura de pdfs.py em vez de busca de campo).
"""

import documentos

ASSINATURAS_PREFIXO = "ASSINATURAS:"
ITEM_ASSINATURA_PADRAO = f"{ASSINATURAS_PREFIXO} Conferir assinaturas de todos os documentos"

FLUXOS = {
    "FQ412-043": {
        "nome": "Liberação de Acordo de Compra do Contrato",
        # Conferir se foi criada ordem de compra no tipo padrão
        "documentos": [
            "Contrato / Aditivo",
            "ACC Master",
            "FQ415-075",
            "FQ412-034",
            "Nota Técnica",
            "Solicitação de Entrega",
        ],
    },
    "FQ415-031": {
        "nome": "OC Padrão – Com Contrato",
        "documentos": [
            "Contrato / Aditivo",
            "Projeto Básico",
            "Nota Técnica",
            "FQ415-075",
        ],
    },
    "FQ412-033": {
        "nome": "Liberação de Acordo de Compra em Aberto",
        # Conferir vigência do contrato no SISCON
        "documentos": [
            "Solicitação de Entrega",
            "FQ415-075",
            "Nota Técnica",
            "FQ412-034",
            "FQ412-035",
        ],
    },
    "FQ415-055": {
        "nome": "Acordo de Compra em Aberto",
        "documentos": [
            "Contrato / Aditivo",
            "Nota Técnica",
            "Projeto Básico",
            "FQ415-075",
        ],
    },
    "FQ415-084": {
        "nome": "Acordo de Compra do Contrato",
        "documentos": [
            "Contrato / Aditivo",
            "Nota Técnica",
            "FQ415-075",
        ],
    },
    "FQ412-069": {
        "nome": "Condomínio",
        "documentos": [
            "Contrato / Aditivo",
            "Ata do Condomínio",
            "Boleto do Condomínio",
            "Nota Técnica",
            "Documento de Referência da Área",
        ],
    },
    "FQ412-067": {
            "nome": "Licitação",
            "documentos": [
                "Contrato / Aditivo",
                "Nota Técnica",
                "Projeto Básico",
            ],
        },
}


def listar_fluxos() -> list:
    """Devolve [(id_fluxo, nome_fluxo), ...] na ordem de definição."""
    return [(fid, dados["nome"]) for fid, dados in FLUXOS.items()]


def documentos_do_fluxo(id_fluxo: str) -> list:
    """Devolve a lista de nomes de documentos de um fluxo, ou [] se o
    fluxo não existir."""
    fluxo = FLUXOS.get(id_fluxo)
    return fluxo["documentos"] if fluxo else []


def campos_do_documento(id_fluxo: str, tipo_documento: str) -> list:
    """Devolve a lista de campos/itens de checklist de um documento dentro
    de um fluxo específico.

    Os campos vêm de documentos.py. Se o documento não fizer parte do
    fluxo, devolve [].
    """
    if tipo_documento not in documentos_do_fluxo(id_fluxo):
        return []

    campos = list(documentos.campos_do_documento(tipo_documento))
    if ITEM_ASSINATURA_PADRAO not in campos:
        campos.append(ITEM_ASSINATURA_PADRAO)

    return campos


def eh_item_assinatura(campo: str) -> bool:
    """True se o item do checklist for uma conferência de assinatura, e não
    um campo a extrair do documento."""
    return campo.startswith(ASSINATURAS_PREFIXO)


def descricao_item_assinatura(campo: str) -> str:
    """Remove o prefixo especial e devolve só a descrição legível do item."""
    return campo[len(ASSINATURAS_PREFIXO):].strip()