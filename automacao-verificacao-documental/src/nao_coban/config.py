"""
Catálogo de insumos para Cadastro/Atualização de Fornecedor "comum" (fora do
fluxo de Credenciamento de Correspondente Bancário - COBAN).

TIPOS_FORNECEDOR espelha (de forma agrupada) as opções que o solicitante
escolhe no Supravizio ao abrir o chamado. A maioria colapsa em só dois
conjuntos de insumos (Pessoa Jurídica / Pessoa Física).

Documentos que aparecem nos pacotes mas NÃO são insumo do TAB415-004
são classificados como  DOCUMENTO_FORA_DE_ESCOPO: reconhecidos para não colidir
por engano com outro tipo, mas não entram em nenhuma validação/pendência.

SUPPORTED_EXTENSIONS e DEFAULT_MIN_OCR_CONFIDENCE (usados por este fluxo e
pelo COBAN) estão em src.shared.constants; não são repetidos aqui.
"""

CNPJ = {
    "label": "Comprovante de Inscrição e Situação Cadastral (CNPJ)",
    "filename_patterns": [r"\bcnpj\b"],
    "content_keywords": [
        "comprovante de inscri",
        "situa",
        "cadastral",
        "matriz",
        "receita federal",
    ],
    "min_ocr_confidence": 80,
}

CONTRATO_SOCIAL = {
    "label": "Contrato Social / Alteração Contratual / CCMEI / Documento Equivalente",
    "filename_patterns": [
        r"contrato\s*social",
        r"altera[cç][aã]o\s*contratual",
        r"\bccmei\b",
        r"\balt\s*\d+",
        r"ata\s*de\s*posse",
        r"ata\s*de\s*elei[cç][aã]o",
        r"altera[cç][aã]o\s*(de\s*)?estatuto",
        r"ata\s*de\s*constitui[cç][aã]o",
    ],
    "content_keywords": [
        "contrato social",
        "altera",
        "clausula",
        "cláusula",
        "microempreendedor individual",
        "junta comercial",
        "ata de posse",
        "diretoria",
        "entidade sindical",
        "estatuto",
    ],
    "min_ocr_confidence": 80,
}

PROCURACAO = {
    # TAB415-004: só é exigida quando quem assina o FQ415-064 não
    # é o representante legal listado no Contrato Social.
    "label": "Procuração",
    "filename_patterns": [r"procura[cç][aã]o"],
    "content_keywords": [
        "procuração",
        "outorgante",
        "outorgado",
        "bastante procurador",
    ],
    "min_ocr_confidence": 55,
}

FORMULARIO_CADASTRO_FORNECEDOR = {
    "label": "Formulário Informações de Cadastro de Fornecedores (FQ415-064)",
    "filename_patterns": [
        r"fq\s*415[\s\-_]*0?64",
        r"cadastro\s*de\s*fornecedor",
        r"cadastro\s*fornecedor",
    ],
    "content_keywords": [
        "informacoes de cadastro de fornecedores",
        "fq415-064",
        "assinatura do representante legal",
        "porte da empresa",
    ],
    "min_ocr_confidence": 80,
}

DOCUMENTO_IDENTIFICACAO = {
    # TAB415-004,: precisa ter foto e estar dentro da validade.
    "label": "Documento de Identificação (RG, CNH, carteira de classe, CTPS ou passaporte)",
    "filename_patterns": [
        r"\brg\b",
        r"\bcnh\b",
        r"carteira\s*de\s*identidade",
        r"passaporte",
        r"\bctps\b",
    ],
    "content_keywords": [
        "carteira nacional de habilitacao",
        "registro geral",
        "carteira de identidade",
        "passaporte",
    ],
    "min_ocr_confidence": 70,
}

CPF = {
    # TAB415-004: dispensável se o número já constar no doc. de identificação.
    "label": "CPF",
    "filename_patterns": [r"\bcpf\b"],
    "content_keywords": ["cadastro de pessoas fisicas", "cpf"],
    "min_ocr_confidence": 70,
}

COMPROVANTE_RESIDENCIA = {
    # TAB415-004: emitido há menos de 90 dias corridos.
    "label": "Comprovante de Residência (emitido há menos de 90 dias)",
    "filename_patterns": [
        r"comprovante\s*.{0,15}resid",
        r"fatura\s*.{0,15}(energia|agua|g[aá]s|telefone|internet)",
    ],
    "content_keywords": ["comprovante de residencia", "fatura", "conta de consumo"],
    "min_ocr_confidence": 60,
}

NOTA_FISCAL_FATURA_BBTS = {
    # TAB415-004: NF/fatura emitida pela concessionária/não
    # fornecedor à BBTS, devidamente atestada pelo gestor.
    "label": "Nota Fiscal ou fatura emitida à BBTS",
    "filename_patterns": [r"nota\s*fiscal", r"\bnf[e]?\b", r"fatura"],
    "content_keywords": ["nota fiscal", "fatura", "bb tecnologia e servicos", "bbts"],
    "min_ocr_confidence": 65,
}

DOCUMENTO_FORA_DE_ESCOPO = {
    "label": "Documento fora do escopo desta validação (não é insumo do TAB415-004)",
    "filename_patterns": [
        r"parecer.*coban",
        r"relat[oó]rio.*ditri",
        r"relatorio[_\s]*288",
        r"fq\s*415[\s\-_]*0?38",
        r"especifica.{0,80}cnica",
        r"projeto.{0,20}b[aá]sico",
        r"\bpb-?e\b",
        # Declaração de opção pela desoneração da folha (CPRB) só existe
        # no pacote quando o fornecedor marcou "SIM" nesse campo do
        # FQ415-064; não é insumo obrigatório do TAB415-004. Sem isso, o
        # texto da declaração contém o CPF de quem assina e acaba
        # confundindo a classificação por conteúdo com o tipo CPF.
        r"\bcprb\b",
        r"desonera.{0,20}folha",
        r"declara.{0,20}(cprb|previdenci)",
    ],
    "content_keywords": [],
    "min_ocr_confidence": 0,
    "fora_de_escopo": True,
}

SIMPLES_NACIONAL = {
    # Não é insumo listado na TAB415-004 pro Cadastro/Atualização. Quando
    # aparece no pacote é evidência que o solicitante anexou por conta
    # própria. Reconhecido mas nunca obrigatório,
    # só pra não colidir com CNPJ por engano na classificação.
    "label": "Consulta de Optante do Simples Nacional (informativo, não é insumo do TAB415-004)",
    "filename_patterns": [r"optante", r"simples\s*nacional", r"consultaoptantes"],
    "content_keywords": ["simples nacional", "optante", "simei"],
    "min_ocr_confidence": 70,
}


def _req(base_cfg: dict, required: bool = True) -> dict:
    """Copia um doc type atômico e define o 'required' pra este tipo de
    fornecedor específico (o mesmo doc type pode ser obrigatório num tipo e
    opcional em outro, ex.: Procuração)."""
    return {**base_cfg, "required": required}


# --- Insumos por tipo de pessoa (TAB415-004, seção "Cadastro de Fornecedor") ----

_INSUMOS_PJ = {
    "CNPJ": _req(CNPJ),
    "CONTRATO_SOCIAL": _req(CONTRATO_SOCIAL),
    "PROCURACAO": _req(PROCURACAO),
    "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
    "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
    "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
}

_INSUMOS_PF = {
    "DOCUMENTO_IDENTIFICACAO": _req(DOCUMENTO_IDENTIFICACAO),
    "CPF": _req(CPF),
    "COMPROVANTE_RESIDENCIA": _req(COMPROVANTE_RESIDENCIA),
    "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
    "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
    "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
}

# --- Os 5 tipos de "Cadastro de Não Fornecedor" (TAB415-004) ---------------

_INSUMOS_OBRIGACOES_JUDICIAIS_PJ = {
    "CNPJ": _req(CNPJ),
    "CONTRATO_SOCIAL": _req(CONTRATO_SOCIAL),
    "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
    "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
    "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
}

_INSUMOS_OBRIGACOES_JUDICIAIS_PF = {
    "DOCUMENTO_IDENTIFICACAO": _req(
        DOCUMENTO_IDENTIFICACAO, required=False
    ),  # "se houver"
    "COMPROVANTE_RESIDENCIA": _req(COMPROVANTE_RESIDENCIA),
    "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
    "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
    "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
}

TIPOS_FORNECEDOR = {
    "FORNECEDOR_CONTRATADO": {
        # Cobre tanto "Cadastro de Fornecedor" (novo) quanto "Atualização
        # Completa e/ou Reativação de Cadastro"
        "label": "Fornecedor Contratado / Atualização Completa (Prestador de Serviços, Aquisição de Bens, Locação de Imóveis, Outros)",
        "insumos_pj": _INSUMOS_PJ,
        "insumos_pf": _INSUMOS_PF,
    },
    # --- Sub-tipos de "Atualização do Cadastro de Fornecedores"
    "INCLUSAO_ATUALIZACAO_TIPO_CONTRIBUINTE": {
        "label": "Inclusão/Atualização de Tipo de Contribuinte",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "ATUALIZACAO_CONTATO_ENDERECO_RAZAO_SOCIAL": {
        "label": "Atualização de Contato, Endereço e/ou Razão Social",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(
                FORMULARIO_CADASTRO_FORNECEDOR, required=False
            ),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "ATUALIZACAO_DADOS_BANCARIOS": {
        "label": "Atualização de Dados Bancários",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "CADASTRO_FILIAL": {
        "label": "Cadastro de Filial",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "CONTRATO_SOCIAL": _req(CONTRATO_SOCIAL, required=False),
            "PROCURACAO": _req(PROCURACAO, required=False),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "CONCESSIONARIA_SERVICO_PUBLICO": {
        "label": "Concessionária de Serviços Públicos",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "NOTA_FISCAL_FATURA_BBTS": _req(NOTA_FISCAL_FATURA_BBTS),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "FUNDOS_ASSISTENCIAIS": {
        "label": "Fundos Assistenciais",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "TRANSPORTADOR_TERCEIROS": {
        "label": "Transportador Contratado por Terceiros / Fornecedor emitente de NF de Remessa sem pagamento",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "NOTA_FISCAL_FATURA_BBTS": _req(NOTA_FISCAL_FATURA_BBTS),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
    "OBRIGACOES_JUDICIAIS": {
        "label": "Obrigações Judiciais",
        "insumos_pj": _INSUMOS_OBRIGACOES_JUDICIAIS_PJ,
        "insumos_pf": _INSUMOS_OBRIGACOES_JUDICIAIS_PF,
    },
    "OBRIGACOES_FISCAIS": {
        "label": "Obrigações Fiscais",
        "insumos": {
            "CNPJ": _req(CNPJ),
            "NOTA_FISCAL_FATURA_BBTS": _req(NOTA_FISCAL_FATURA_BBTS),
            "FORMULARIO_CADASTRO_FORNECEDOR": _req(FORMULARIO_CADASTRO_FORNECEDOR),
            "DOCUMENTO_FORA_DE_ESCOPO": _req(DOCUMENTO_FORA_DE_ESCOPO, required=False),
            "SIMPLES_NACIONAL": _req(SIMPLES_NACIONAL, required=False),
        },
    },
}
