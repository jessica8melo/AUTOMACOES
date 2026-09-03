"""
Classificação de documentos da OS.

  1) Nome do arquivo (regex) -> classificação imediata, sem custo de OCR.
  2) Se o nome não bater com nada (ou 'empatar' em mais de um tipo), usa o
     texto já extraído (camada de texto do PDF ou OCR) para 'desempatar' via
     contagem de palavras-chave.

Usada pelos dois fluxos (COBAN e Não COBAN): por isso `doc_types` é sempre
obrigatório aqui; quem chama passa o catálogo do próprio fluxo
(src.coban.config.DOCUMENT_TYPES ou o dict de insumos do
src.nao_coban.config.TIPOS_FORNECEDOR[tipo]). Este módulo não conhece
nenhum catálogo específico.
"""

import re
import unicodedata
from typing import Optional


def _strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)
    )


# Ordem de checagem por nome de arquivo. Usada pelos dois fluxos; só
# entram na comparação os tipos que existirem em `doc_types`.
_FILENAME_CHECK_ORDER = [
    "PARECER_COBAN",
    "FORMULARIO_CREDENCIAMENTO",
    "COMPROVANTE_BANCARIO",
    "CNPJ",
    "CONTRATO_SOCIAL",
    "SIMPLES_NACIONAL",
]


def classify_by_filename(filename: str, doc_types: dict) -> Optional[str]:
    """Retorna a chave em doc_types cujo padrão bate com o nome do arquivo,
    ou None se nenhum padrão bater. Segue _FILENAME_CHECK_ORDER para
    resolver ambiguidades entre tipos.
    """
    name_norm = _strip_accents(filename).lower()
    check_order = _FILENAME_CHECK_ORDER + [
        dt for dt in doc_types if dt not in _FILENAME_CHECK_ORDER
    ]
    for doc_type in check_order:
        cfg = doc_types.get(doc_type)
        if not cfg:
            continue
        for pattern in cfg["filename_patterns"]:
            if re.search(pattern, name_norm):
                return doc_type
    return None


def classify_by_content(text: str, doc_types: dict) -> Optional[str]:
    """Desempate por conteúdo: conta ocorrências de content_keywords de cada
    tipo de documento e retorna o tipo com mais acertos (mínimo de 1 acerto).
    """
    if not text:
        return None
    text_norm = _strip_accents(text).lower()

    scores = {}
    for doc_type, cfg in doc_types.items():
        score = 0
        for kw in cfg["content_keywords"]:
            kw_norm = _strip_accents(kw).lower()
            if kw_norm in text_norm:
                score += 1
        if score:
            scores[doc_type] = score

    if not scores:
        return None
    return max(scores, key=scores.get)


def classify_document(filename: str, text: str, doc_types: dict) -> str:
    """Classificação final: tenta nome do arquivo primeiro; se falhar, usa
    conteúdo; se ambos falharem, retorna 'DESCONHECIDO'.
    """
    by_name = classify_by_filename(filename, doc_types)
    if by_name:
        return by_name

    by_content = classify_by_content(text, doc_types)
    if by_content:
        return by_content

    return "DESCONHECIDO"
