"""
Peças genéricas de validação/comparação de texto, usadas tanto pelos
validadores compartilhados (src.shared.common_validators) quanto pelos
validadores específicos de cada fluxo (src.coban.validators,
src.nao_coban.validators): a dataclass de resultado de validação, extração/
normalização de CNPJ, e comparação "tolerante a ruído de OCR" de nomes.
"""

import difflib
import re
import unicodedata
from dataclasses import dataclass, field

CNPJ_RE = re.compile(
    r"\d{2}[^\d]{0,2}\d{3}[^\d]{0,2}\d{3}[^\d]{0,2}\d{4}[^\d]{0,2}\d{2}"
)
DATE_RE = re.compile(r"\d{2}/\d{2}/\d{4}")


def _normalizar_nome(s: str) -> str:
    """Maiúsculas, sem acento, sem pontuação: pra comparar
    com variações de grafia/OCR do mesmo nome.
    """
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip().upper()


def nome_empresarial_confere(nome_referencia: str, texto_documento: str) -> bool:
    ref = _normalizar_nome(nome_referencia)
    if not ref:
        return False
    texto_norm = _normalizar_nome(texto_documento)

    if ref in texto_norm:
        return True

    ignorar = {"LTDA", "ME", "EPP", "EIRELI", "DE", "DA", "DO", "E"}
    palavras_ref = [p for p in ref.split() if len(p) >= 3 and p not in ignorar]
    if not palavras_ref:
        return False
    encontradas = sum(1 for p in palavras_ref if p in texto_norm)
    return encontradas / len(palavras_ref) >= 0.8


def nomes_correspondem(nome_a: str, nome_b: str, limiar: float = 0.95) -> bool:
    a = _normalizar_nome(nome_a)
    b = _normalizar_nome(nome_b)

    if not a or not b:
        return False
    if a in b or b in a:
        return True

    return difflib.SequenceMatcher(None, a, b).ratio() >= limiar


def _normalize_cnpj(raw: str) -> str:
    return re.sub(r"\D", "", raw)


def find_cnpj(text: str) -> "str | None":
    m = CNPJ_RE.search(text or "")
    return _normalize_cnpj(m.group()) if m else None


def find_all_cnpjs(text: str) -> list:
    """Todos os números no formato de CNPJ presentes no texto (não só o
    primeiro): um Contrato Social pode ter mais de um CNPJ (o do
    cartório que registrou o documento, o de outra parte citada etc.), e
    nem sempre o primeiro que aparece é o da própria empresa."""
    return [_normalize_cnpj(m.group()) for m in CNPJ_RE.finditer(text or "")]


@dataclass
class DocumentValidation:
    doc_type: str
    filename: str
    valido: bool
    campos: dict[str, str] = field(default_factory=dict)
    pendencias: list[str] = field(default_factory=list)


def cross_validate_cnpj(validations: list[DocumentValidation]) -> list[str]:
    """Confere se o CNPJ é o mesmo em todos os documentos que reportaram um."""
    cnpjs = {v.campos.get("cnpj") for v in validations if v.campos.get("cnpj")}
    if len(cnpjs) > 1:
        return [
            f"CNPJs divergentes entre documentos do pacote: {', '.join(sorted(cnpjs))}."
        ]
    return []
