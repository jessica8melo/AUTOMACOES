# -*- coding: utf-8 -*-
"""
Tabela de Localidade (padrão de referência para buscas).
Gerado a partir da planilha original fornecida pelo usuário.

Cada item representa uma linha da tabela original, com as colunas:
    escritorio  -> Escritório
    cat         -> CAT
    sigla       -> Sigla
    org         -> ORG.
    vencimento  -> Vencimento
    status      -> Status
    observacao  -> Observação
"""

TABELA_LOCALIDADE = [
    {"escritorio": "Contmax", "cat": "Rio de Janeiro",          "sigla": "RIO",   "org": "MMT", "vencimento": "3º dia útil", "status": "ok",  "observacao": "ISS Retido do mês"},
    {"escritorio": "Contmax", "cat": "Rio de Janeiro",          "sigla": "RIO",   "org": "MMT", "vencimento": "3º dia útil", "status": "S/M", "observacao": "ISS IMPORTAÇÃO Ref. a invoice - Sem movimento em 01/2025 conforme afirmação de maria.marques@bbts.com.br - 05/02/2025"},
    {"escritorio": "Contmax", "cat": "Estoque Central/RJ",      "sigla": "MCE",   "org": "MCE", "vencimento": "3º dia útil", "status": "ok",  "observacao": ""},
    {"escritorio": "Contmax", "cat": "Carioca/RJ",              "sigla": "CAR",   "org": "MRC", "vencimento": "3º dia útil", "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Salvador",                "sigla": "SAL",   "org": "MSA", "vencimento": "5",           "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Salvador 2",              "sigla": "SAL 2", "org": "STA", "vencimento": "5",           "status": "ok",  "observacao": ""},
    {"escritorio": "Ditri",   "cat": "Lauro de Freitas-BA",     "sigla": "LAU",   "org": "LAU", "vencimento": "5",           "status": "S/M", "observacao": "Sem movimento em 01/2025, conforme Relatório NFS Entradas no RJ/ERP"},
    {"escritorio": "Véritas", "cat": "Belo Horizonte",          "sigla": "BHZ",   "org": "MBH", "vencimento": "8",           "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Belém",                   "sigla": "BEM",   "org": "MBE", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Micael",  "cat": "Campinas",                "sigla": "CAM",   "org": "MCA", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "NCI",     "cat": "Fortaleza",               "sigla": "FOR",   "org": "MFO", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Contmax", "cat": "Goiânia",                 "sigla": "GOI",   "org": "MGO", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Contmax", "cat": "Goiânia 2",               "sigla": "GOI 2", "org": "CIG", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "NCI",     "cat": "João Pessoa",             "sigla": "JPA",   "org": "MJP", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Manaus",                  "sigla": "MAN",   "org": "MMA", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "NCI",     "cat": "Natal",                   "sigla": "NAT",   "org": "MNA", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Contmax", "cat": "Piraí",                   "sigla": "PIR",   "org": "MPI", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "NCI",     "cat": "Recife",                  "sigla": "REC",   "org": "MRE", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Micael",  "cat": "Porto Alegre",            "sigla": "POA",   "org": "MPA", "vencimento": "10",          "status": "ok",  "observacao": "Guia ISS retido junto com o Próprio"},
    {"escritorio": "Véritas", "cat": "São Paulo",               "sigla": "SPO",   "org": "MSP", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Teresina",                "sigla": "TER",   "org": "MTE", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Contmax", "cat": "Vitória",                 "sigla": "VIT",   "org": "MVI", "vencimento": "10",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Guarulhos",               "sigla": "GRU",   "org": "MBI", "vencimento": "12",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "São Luís",                "sigla": "SLU",   "org": "MSL", "vencimento": "12",          "status": "ok",  "observacao": ""},
    {"escritorio": "Micael",  "cat": "Bauru",                   "sigla": "BAU",   "org": "MBA", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Véritas", "cat": "Cascavel",                "sigla": "CAS",   "org": "MCC", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Micael",  "cat": "Campo Grande",            "sigla": "CGR",   "org": "MCG", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Contmax", "cat": "Florianópolis",           "sigla": "FLO",   "org": "MFL", "vencimento": "15",          "status": "ok",  "observacao": "Antecipação para dia 14/02/2025"},
    {"escritorio": "Micael",  "cat": "Joinville",               "sigla": "JOI",   "org": "MJO", "vencimento": "15",          "status": "ok",  "observacao": "Antecipação para dia 14/02/2025"},
    {"escritorio": "Micael",  "cat": "Londrina",                "sigla": "LON",   "org": "MLO", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Micael",  "cat": "Passo Fundo",             "sigla": "PAF",   "org": "MPF", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Véritas", "cat": "Palmas",                  "sigla": "PAM",   "org": "MPM", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Véritas", "cat": "Porto Velho",             "sigla": "POV",   "org": "MPV", "vencimento": "15",          "status": "ok",  "observacao": "Antecipação para dia 14/02/2025"},
    {"escritorio": "Véritas", "cat": "Ribeirão Preto",          "sigla": "RIP",   "org": "MRP", "vencimento": "15",          "status": "ok",  "observacao": "Guia 17/02/2025 mesmo mês"},
    {"escritorio": "Véritas", "cat": "Uberlândia",              "sigla": "UBE",   "org": "MUB", "vencimento": "15",          "status": "ok",  "observacao": "Antecipação para dia 14/02/2025"},
    {"escritorio": "Véritas", "cat": "Brasília",                "sigla": "BRA",   "org": "MBR", "vencimento": "20",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Cuiabá",                  "sigla": "CUI",   "org": "MCB", "vencimento": "20",          "status": "ok",  "observacao": ""},
    {"escritorio": "Véritas", "cat": "Curitiba",                "sigla": "CUR",   "org": "MCR", "vencimento": "20",          "status": "ok",  "observacao": ""},
    {"escritorio": "NCI",     "cat": "Maceió",                  "sigla": "MAC",   "org": "MMC", "vencimento": "20",          "status": "ok",  "observacao": ""},
    {"escritorio": "Micael",  "cat": "Osasco",                  "sigla": "OSA",   "org": "OPL", "vencimento": None,          "status": None,  "observacao": ""},
    {"escritorio": None,      "cat": "Barueri",                 "sigla": None,    "org": "BAI", "vencimento": None,          "status": None,  "observacao": ""},
]


def buscar_por_sigla(sigla):
    """Retorna o registro cuja sigla corresponde (case-insensitive), ou None."""
    sigla = sigla.strip().upper()
    for item in TABELA_LOCALIDADE:
        if item["sigla"] and item["sigla"].upper() == sigla:
            return item
    return None


def buscar_por_cat(cat):
    """Retorna o registro cujo CAT (nome da localidade) corresponde (case-insensitive), ou None."""
    cat = cat.strip().lower()
    for item in TABELA_LOCALIDADE:
        if item["cat"] and item["cat"].lower() == cat:
            return item
    return None


def buscar_por_org(org):
    """Retorna o registro cujo ORG corresponde (case-insensitive), ou None."""
    org = org.strip().upper()
    for item in TABELA_LOCALIDADE:
        if item["org"] and item["org"].upper() == org:
            return item
    return None


if __name__ == "__main__":
    # Exemplo de uso
    print(f"Total de registros: {len(TABELA_LOCALIDADE)}")
    print(buscar_por_sigla("RIO"))