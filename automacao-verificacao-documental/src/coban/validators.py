"""
Regras de validação específicas do fluxo COBAN (Credenciamento de
Correspondente Bancário). Cada validador recebe o texto já extraído 
(src.shared.extractor) e devolve um DocumentValidation com:
  - valido: passou nas regras mínimas?
  - campos: dados estruturados extraídos (CNPJ, situação, datas...)
  - pendencias: lista de problemas encontrados (para o relatório e e-mail)

CNPJ, Contrato Social, Simples Nacional e Procuração são validados pelas
mesmas funções do fluxo Não COBAN; ver src.shared.common_validators.
"""

import re

from src.shared.common_validators import (
    ASSINANTE_D4SIGN_RE,
    validate_cnpj_doc,
    validate_contrato_social,
    validate_procuracao,
    validate_simples_nacional,
)
from src.shared.validation import DATE_RE, DocumentValidation, find_cnpj

NOME_SOCIO_FORM_RE = re.compile(r"Nome\s+S[oó]cio\s*\d*\s*:\s*([^\n\r]+)", re.I)


def buscar_nomes_socios_formulario(text: str) -> "list[str]":
    """Fonte confiável do(s) nome(s) de sócio: o próprio Formulário de
    Credenciamento tem um campo rotulado 'Nome Sócio 1:', 'Nome Sócio 2:'
    """
    return [m.group(1).strip() for m in NOME_SOCIO_FORM_RE.finditer(text or "")]


def validate_formulario_credenciamento(text: str, filename: str) -> DocumentValidation:

    pendencias = []
    cnpj = find_cnpj(text)
    if not cnpj:
        pendencias.append("CNPJ não localizado no Formulário de Credenciamento.")

    assinado = bool(re.search(r"d4sign", text, re.I))
    if not assinado:
        pendencias.append(
            "Não foi possível confirmar assinatura via D4Sign no documento."
        )

    nomes_socios = buscar_nomes_socios_formulario(text)

    if not nomes_socios:
        pendencias.append(
            "Campo 'Nome Sócio' não localizado no Formulário. Verificar manualmente."
        )

    nomes_assinantes_d4sign = [
        m.group(1).strip()
        for m in ASSINANTE_D4SIGN_RE.finditer(text)
        if not m.group(2).lower().endswith("@bbts.com.br")
    ]

    return DocumentValidation(
        doc_type="FORMULARIO_CREDENCIAMENTO",
        filename=filename,
        valido=not pendencias,
        campos={
            "cnpj": cnpj or "",
            "nomes_socios": ";".join(nomes_socios),
            "nomes_assinantes_d4sign": ";".join(nomes_assinantes_d4sign),
        },
        pendencias=pendencias,
    )


def validate_parecer_coban(text: str, filename: str) -> DocumentValidation:
    """O 'Parecer COBAN' descreve retenção de impostos
    (INSS/IRRF/PIS/COFINS/CSLL/ISS) e o enquadramento do fornecedor
    (regime normal / Simples Nacional).
    """
    pendencias = []

    responsavel_match = re.search(
        r"Respons[aá]vel pela an[aá]lise:\s*([^\n\r]+)", text, re.I
    )
    responsavel = responsavel_match.group(1).strip() if responsavel_match else ""
    if not responsavel:
        pendencias.append("Responsável pela análise não identificado no Parecer COBAN.")

    data_analise = ""
    if responsavel_match:
        trecho = text[responsavel_match.end() : responsavel_match.end() + 200]
        data_match = DATE_RE.search(trecho)
        data_analise = data_match.group() if data_match else ""
    if not data_analise:
        pendencias.append("Data da análise não identificada no Parecer COBAN.")

    tipo_contribuinte_match = re.search(
        r"Tipo de Contribuinte ser[aá]:\s*([^\n\r]+)", text, re.I
    )
    tipo_contribuinte = (
        tipo_contribuinte_match.group(1).strip() if tipo_contribuinte_match else ""
    )

    optante_simples = bool(
        re.search(r"optante\s+pelo\s+simples\s+nacional", text, re.I)
    )
    nao_optante_simples = bool(re.search(r"n[ãa]o\s+(?:é\s+)?optante", text, re.I))

    return DocumentValidation(
        doc_type="PARECER_COBAN",
        filename=filename,
        valido=not pendencias,
        campos={
            "responsavel_analise": responsavel,
            "data_analise": data_analise,
            "tipo_contribuinte": tipo_contribuinte,
            "optante_simples": str(optante_simples and not nao_optante_simples),
        },
        pendencias=pendencias,
    )


AGENCIA_CONTA_LINHA_RE = re.compile(
    r"ag[eê]ncia\s+conta\s+([\d.\-]+)\s+([\d.\-]+)", re.I
)
AGENCIA_RE = re.compile(
    r"(?:ag[eê]ncia|ag\.)\s*(?:n?[ºo°]?)?\s*[:.]?\s*([\d\-\.]{2,15})", re.I
)
CONTA_RE = re.compile(
    r"(?:conta(?:\s+corrente)?|cc\.?|c/c)\s*(?:n?[ºo°]?)?\s*[:.]?\s*([\d\-\.]{3,20})",
    re.I,
)
BANCO_RE = re.compile(r"banco[:\s]+([A-Za-zÀ-ÿ0-9 \.]{3,40})", re.I)
AG_CC_RE = re.compile(r"ag\.?\s*([\d\-\.]+).*?cc\.?\s*([\d\-\.]+)", re.I | re.S)
AGENCIA_CONTA_TABELA_RE = re.compile(
    r"ag[eê]ncia\s*([\d.\-]+)\s*conta\s*([\d.\-]+)", re.I | re.S
)


def _find_value_below_label(label_re: re.Pattern, text: str, window: int = 100) -> str:
    m = label_re.search(text or "")
    if not m:
        return ""

    trecho = text[m.end() : m.end() + window]

    valor = re.search(r"\b\d[\d\.\-]{1,20}\b", trecho)
    return valor.group() if valor else ""


def validate_comprovante_bancario(text: str, filename: str) -> DocumentValidation:
    pendencias = []

    agencia = ""
    conta = ""

    linha_match = AGENCIA_CONTA_LINHA_RE.search(text)

    if linha_match:
        agencia = linha_match.group(1).strip()
        conta = linha_match.group(2).strip()

    # Formato tabela
    elif table_match := AGENCIA_CONTA_TABELA_RE.search(text):
        agencia = table_match.group(1).strip()
        conta = table_match.group(2).strip()

    # Formato Ag. xxxx Cc. yyyy
    elif pair_match := AG_CC_RE.search(text):
        agencia = pair_match.group(1).strip()
        conta = pair_match.group(2).strip()

    # Busca individual
    if not agencia:
        agencia_match = AGENCIA_RE.search(text)
        if agencia_match:
            agencia = agencia_match.group(1).strip()
        else:
            agencia = _find_value_below_label(
                re.compile(r"(?:ag[eê]ncia|ag\.)", re.I), text
            )

    if not conta:
        conta_match = CONTA_RE.search(text)
        if conta_match:
            conta = conta_match.group(1).strip()
        else:
            conta = _find_value_below_label(
                re.compile(r"(?:conta(?:\s+corrente)?|cc\.?|c/c)", re.I), text
            )

    banco_match = BANCO_RE.search(text)
    banco = banco_match.group(1).strip() if banco_match else ""

    if not agencia:
        pendencias.append(
            "Agência não localizada no comprovante bancário. Verificar manualmente."
        )
    if not conta:
        pendencias.append(
            "Conta não localizada no comprovante bancário. Verificar manualmente."
        )

    return DocumentValidation(
        doc_type="COMPROVANTE_BANCARIO",
        filename=filename,
        valido=not pendencias,
        campos={"banco": banco, "agencia": agencia, "conta": conta},
        pendencias=pendencias,
    )


VALIDATORS = {
    "CNPJ": validate_cnpj_doc,
    "CONTRATO_SOCIAL": validate_contrato_social,
    "SIMPLES_NACIONAL": validate_simples_nacional,
    "FORMULARIO_CREDENCIAMENTO": validate_formulario_credenciamento,
    "PARECER_COBAN": validate_parecer_coban,
    "COMPROVANTE_BANCARIO": validate_comprovante_bancario,
    "PROCURACAO": validate_procuracao,
}


def validate_document(doc_type: str, text: str, filename: str) -> DocumentValidation:
    validator = VALIDATORS.get(doc_type)
    if not validator:
        return DocumentValidation(
            doc_type=doc_type or "DESCONHECIDO",
            filename=filename,
            valido=False,
            pendencias=[
                "Tipo de documento não reconhecido. Revisão manual necessária."
            ],
        )
    return validator(text, filename)
