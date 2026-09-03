"""
Catálogo dos tipos de documento esperados em uma OS de Credenciamento de
Correspondente Bancário (fluxo COBAN), com as regras usadas por
src.shared.classifier e src.coban.validators.

Cada entrada define:
- filename_patterns: regex (case-insensitive) aplicados ao NOME do arquivo.
  Usados como primeiro sinal de classificação (mais rápido que OCR).
- content_keywords: termos que devem aparecer no texto extraído (OCR ou
  camada de texto do PDF) para confirmar a classificação e servir de
  segundo sinal quando o nome do arquivo é ambíguo ou genérico.
- required: se True, a ausência do documento no pacote é reportada como
  pendência (o resultado final do OS depende de todos os "required" = True
  estarem presentes e válidos).

SUPPORTED_EXTENSIONS e DEFAULT_MIN_OCR_CONFIDENCE (usados por este fluxo e
pelo Não COBAN) em src.shared.constants; não são repetidos aqui.
"""

DOCUMENT_TYPES = {
    "CNPJ": {
        "label": "Comprovante de Inscrição e Situação Cadastral (CNPJ)",
        "filename_patterns": [r"\bcnpj\b"],
        "content_keywords": [
            "comprovante de inscri",
            "situa",
            "cadastral",
            "matriz",
            "receita federal",
        ],
        "required": True,
        "min_ocr_confidence": 80,
    },
    "CONTRATO_SOCIAL": {
        "label": "Contrato Social / Alteração Contratual / CCMEI",
        "filename_patterns": [
            r"contrato\s*social",
            r"altera[cç][aã]o\s*contratual",
            r"\bccmei\b",
            r"\balt\s*\d+",
            r"ata\s*de\s*posse",
            r"ata\s*de\s*elei[cç][aã]o",
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
        ],
        "required": True,
        "min_ocr_confidence": 80,
    },
    "SIMPLES_NACIONAL": {
        "label": "Consulta de Optante do Simples Nacional",
        "filename_patterns": [
            r"optante",
            r"simples\s*nacional",
            r"consultaoptantes",
            r"\bdb\b",
        ],
        "content_keywords": [
            "simples nacional",
            "optante",
            "simei",
        ],
        "required": True,
        "min_ocr_confidence": 70,
    },
    "FORMULARIO_CREDENCIAMENTO": {
        "label": "Formulário de Credenciamento (assinado via D4Sign)",
        "filename_patterns": [r"formul[aá]rio.*credenciamento", r"d4sign"],
        "content_keywords": [
            "cadastro de correspondentes",
            "fq2000",
            "informações cadastrais",
            "d4sign",
        ],
        "required": True,
        "min_ocr_confidence": 85,
    },
    "PARECER_COBAN": {
        "label": "Parecer COBAN",
        "filename_patterns": [r"parecer.*coban"],
        "content_keywords": ["parecer", "coban"],
        "required": True,
        "min_ocr_confidence": 70,
    },
    "COMPROVANTE_BANCARIO": {
        "label": "Comprovante Bancário (conta PJ)",
        "filename_patterns": [
            r"comprovante\s*banc",
            r"conta\s*pj",
            r"\bbb\b\.(jpg|jpeg|png)$",
            r"\bdb\b",
            r"comprovante",
        ],
        "content_keywords": [
            "agência",
            "agencia",
            "conta corrente",
            "banco do brasil",
            "titular",
        ],
        "required": False,
        "min_ocr_confidence": 60,
    },
    "PROCURACAO": {
        # TAB415-004, insumos de Pessoa Jurídica: "Procuração outorgando
        # poderes para assinar acordos/contratos e representar o fornecedor".
        # Só é obrigatória quando quem assina o Formulário NÃO é o sócio listado
        "label": "Procuração",
        "filename_patterns": [r"procura[cç][aã]o"],
        "content_keywords": [
            "procuração",
            "outorgante",
            "outorgado",
            "bastante procurador",
        ],
        "required": False,
        "min_ocr_confidence": 55,
    },
}
