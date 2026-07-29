#!/usr/bin/env python3
"""
Documentos reconhecidos e campos a extrair de cada um.

Este módulo é o ponto único de configuração da etapa de EXTRAÇÃO: para
cada tipo de documento (Contrato, FQ415-075, Nota Técnica, Projeto
Básico, Solicitação de Entrega), lista os campos que devem ser
localizados dentro do arquivo.

Por enquanto o projeto foca só nisso — identificar o documento e extrair
seus campos — sem amarrar isso a um fluxo/checklist específico (isso
fica para depois, em fluxos.py).

Para adicionar, remover ou renomear campos de um documento, edite
livremente o dicionário abaixo.
"""

DOCUMENTOS = {
    "Contrato / Aditivo": [
        "DGCO",
        "OC Master",
        "Tipo de Contratação",
        "Contratada",                  # Fornecedor
        "CNPJ",
        "Quantidade",
        "Preço unitário",
        "Valor total",
        "Data do contrato",         
    ],
    "FQ415-075": [
        "Código do Item",
        "Descrição do Item",
        "Unidade de Medida",            # UDM
        "Quantidade do Item",           # Quantidade
        "Valor/Preço Unitário",         # Preço
        "Valor Total",                  # Total
        "Natureza da Transação",        # Natureza da Transação
        "Local para Faturamento",       # Entregar Para/Faturar Para/Modalida
        "UOR",                          # Conta de Débito
        "Conta Contábil",
    ],
    "Nota Técnica": [
        "Número da Nota Técnica",
    ],
    "Projeto Básico": [
        "Codigo do Item",
        "Descrição do Item",
        "Natureza da transação",
        "Natureza contábil",
        "Conta contábil de despesa ou investimento",
        "UOR",
        "CNPJ de faturamento",
        "Condições de Garantia e Assistência Técnica, Manutenção e Suporte Técnico",
        "Condições de Pagamento",       # dias corridos
    ],
    "Solicitação de Entrega": [
        "DGCO",
        "OC Master",
        "Valor total da solicitação",   # Total
        "Endereço",                     # Entregar Para
        "Quantidade",                   # Quantidade
        "Data do fornecimento",         # Prometido
        "Código BBTS",                  # Código
        "Especificação do Bem",         # Nome
    ],
}


def listar_documentos() -> list:
    """Devolve os nomes dos documentos reconhecidos, na ordem de definição."""
    return list(DOCUMENTOS.keys())


def campos_do_documento(tipo_documento: str) -> list:
    """Devolve a lista de campos a extrair de um documento, ou [] se o tipo
    de documento não for reconhecido."""
    return DOCUMENTOS.get(tipo_documento, [])
