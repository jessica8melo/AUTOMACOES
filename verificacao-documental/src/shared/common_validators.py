"""
Validadores usados IDENTICAMENTE pelos dois fluxos (COBAN e Não COBAN):
CNPJ, Contrato Social, Simples Nacional e Procuração são o mesmo documento
com as mesmas regras nos dois; não há motivo pra duplicar. Cada fluxo só
registra estas funções na própria tabela de validadores
(src.coban.validators.VALIDATORS / src.nao_coban.validators.VALIDATORS_NAO_COBAN).

ASSINANTE_D4SIGN_RE também é compartilhada: usada tanto pelo Formulário de
Credenciamento (COBAN) quanto pelo FQ415-064 (Não COBAN) pra achar quem
assinou via D4Sign.
"""

import re
import unicodedata
from datetime import datetime

from src.shared.validation import DocumentValidation, find_cnpj

NOME_EMPRESARIAL_RE = re.compile(r"^NOME EMPRESARIAL\s*\n\s*([^\n\r]+)", re.I | re.M)
NOME_EMPRESARIAL_PROSA_RE = re.compile(
    r"nome\s+empresarial\s+([A-ZÀ-Ú][^.\n]{2,80}?)\s*\.", re.I
)


NOME_EMPRESARIAL_DENOMINADA_RE = re.compile(
    r"denominada\s*[:\s]*[\"'“‘]([^\"'”’]{2,100})[\"'”’]", re.I
)
NOME_EMPRESARIAL_SEDE_RE = re.compile(
    r"sede\s+d[ae]\s+([A-ZÀ-Ü][^(\n]{2,100}?)\s*\(", re.I
)
NOME_EMPRESARIAL_ANTES_CNPJ_RE = re.compile(
    r"([A-ZÀ-Ü][A-ZÀ-Üa-zà-ÿ0-9\.\-]{1,30}(?:\s+[A-ZÀ-Üa-zà-ÿ0-9\.\-]{1,30}){0,5}\s"
    r"(?:LTDA|S\.?A\.?|S/A|EIRELI|ME|EPP))\.?\s+CNPJ",
    re.I,
)


def extrair_nome_empresarial(text: str) -> str:

    text_norm = (
        unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode()
    )

    m = re.search(
        r"NOME\s+EMPRESARIAL\s+(.{1,120}?)\s+(?:TITULO\s+DO\s+ESTABELECIMENTO|CODIGO\s+E\s+DESCRI[CÇ][AÃ]O\s+DA\s+ATIVIDADE)",
        text_norm,
        re.I | re.S,
    )

    if m:
        nome = " ".join(m.group(1).split())
        return nome.strip()

    m = NOME_EMPRESARIAL_PROSA_RE.search(text or "")
    if m:
        return m.group(1).strip()

    m = NOME_EMPRESARIAL_DENOMINADA_RE.search(text or "")
    if m:
        return m.group(1).strip()

    m = NOME_EMPRESARIAL_SEDE_RE.search(text or "")
    if m:
        return m.group(1).strip()

    m = NOME_EMPRESARIAL_ANTES_CNPJ_RE.search(text or "")
    if m:
        return " ".join(m.group(1).split())

    return ""


DATA_EMISSAO_RE = re.compile(r"emitido\s+no\s+dia\s+(\d{2}/\d{2}/\d{4})", re.I)


def validate_cnpj_doc(text: str, filename: str) -> DocumentValidation:
    """Confere CNPJ, situação cadastral, e a atualidade do próprio
    comprovante, que é dada pela linha 'Emitido no dia DD/MM/AAAA...'.
    """
    pendencias = []
    cnpj = find_cnpj(text)

    if not cnpj:
        pendencias.append("CNPJ não localizado no documento.")

    situacao = "ATIVA" if re.search(r"\bativa\b", text, re.I) else None
    if not situacao:
        pendencias.append(
            "Situação cadastral 'ATIVA' não identificada. Verificar manualmente."
        )

    emissao_match = DATA_EMISSAO_RE.search(text)
    data_emissao = emissao_match.group(1) if emissao_match else ""
    ano_emissao = data_emissao.split("/")[-1] if data_emissao else ""

    if not data_emissao:
        pendencias.append(
            "Data de emissão do comprovante não localizada. Verificar manualmente."
        )
    else:
        ano_atual = datetime.now().year

        if int(ano_emissao) < ano_atual:
            pendencias.append(
                f"Comprovante de CNPJ com data de emissão desatualizada (ano {ano_emissao}). Solicitar comprovante mais recente."
            )

    nome_empresarial = extrair_nome_empresarial(text)

    return DocumentValidation(
        doc_type="CNPJ",
        filename=filename,
        valido=not pendencias,
        campos={
            "cnpj": cnpj or "",
            "situacao_cadastral": situacao or "NAO_IDENTIFICADA",
            "data_emissao": data_emissao,
            "ano_emissao": ano_emissao,
            "nome_empresarial": nome_empresarial,
        },
        pendencias=pendencias,
    )


def validate_contrato_social(text: str, filename: str) -> DocumentValidation:
    """Quando não há CNPJ próprio no texto (comum em Alterações
    Contratuais, que só citam o nome da empresa), essa pendência pode ser
    resolvida depois no orchestrator, comparando o nome empresarial
    com o do Cartão CNPJ (documento-base).
    """
    pendencias = []
    cnpj = find_cnpj(text)
    if not cnpj and not re.search(r"microempreendedor individual", text, re.I):
        pendencias.append("CNPJ não localizado no Contrato Social/CCMEI.")

    nome_empresarial = extrair_nome_empresarial(text)

    return DocumentValidation(
        doc_type="CONTRATO_SOCIAL",
        filename=filename,
        valido=not pendencias,
        campos={"cnpj": cnpj or "", "nome_empresarial": nome_empresarial},
        pendencias=pendencias,
    )


def validate_simples_nacional(text: str, filename: str) -> DocumentValidation:

    pendencias = []

    cnpj = find_cnpj(text)
    if not cnpj:
        pendencias.append("CNPJ não localizado na consulta do Simples Nacional.")

    situacao_match = re.search(
        r"situa[cç][aã]o\s+no\s+simples\s+nacional:\s*([^\n\r]+)",
        text,
        re.I,
    )

    situacao = situacao_match.group(1).strip() if situacao_match else ""

    optante = "NÃO OPTANTE" not in situacao.upper() if situacao else None

    if optante is None:
        pendencias.append("Situação no Simples Nacional não localizada.")

    return DocumentValidation(
        doc_type="SIMPLES_NACIONAL",
        filename=filename,
        valido=not pendencias,
        campos={
            "cnpj": cnpj or "",
            "optante_simples": str(optante),
            "situacao_simples": situacao,
        },
        pendencias=pendencias,
    )


ASSINANTE_D4SIGN_RE = re.compile(
    r"([^\n\r]{3,100})\s*\n\s*([\w.\-+]+@[\w.\-]+)\s*\n\s*Assinou", re.I
)


OUTORGANTE_RE = re.compile(r"outorgante[:\s]*([^\n\r,]+)", re.I)
OUTORGADO_RE = re.compile(r"outorgado[:\s]*([^\n\r,]+)", re.I)


def _valor_parece_nome(valor: str) -> bool:
    """OUTORGANTE_RE/OUTORGADO_RE pegam a palavra 'outorgante'
    onde quer que apareça, mesmo dentro de uma frase como (“Outorgante”)
    pode capturar lixo de pontuação em vez de um nome de verdade.
    Exige pelo menos 2 letras seguidas pra aceitar o valor."""
    return bool(re.search(r"[A-Za-zÀ-ÿ]{2,}", valor or ""))


_SECAO_OUTORGADOS_INICIO_RE = re.compile(r"OUTORGADOS?\s*:", re.I)
_SECAO_OUTORGADOS_FIM_RE = re.compile(r"\b(PODERES|RESTRI[CÇ][OÕ]ES)\s*:", re.I)
_NOME_LISTA_NUMERADA_RE = re.compile(r"^\s*\d+\.\s+([A-ZÀ-Ý][^,\n\r]{2,80}),", re.M)


_SECAO_OUTORGADOS_INICIO_ALT_RE = re.compile(
    r"nomeia\s+e\s+constitui\s+(?:seus?|sua)\s+procurador", re.I
)
_NOME_BRASILEIRO_RE = re.compile(r"([A-ZÀ-Ý][A-ZÀ-Ý\s]{2,60}),\s*brasileir[oa]")


def extrair_outorgados(text: str) -> "list[str]":
    """Extrai todos os nomes de outorgados/procuradores de uma Procuração.
    Cobre 3 formatos vistos na prática:
    1) Lista numerada logo após 'OUTORGADOS:' (com ou sem 'GRUPO A'/'GRUPO B').
    2) Um único outorgado citado logo após 'OUTORGADO(A):'.
    3) Nomes sem lista numerada: 'GRUPO 1: FULANO, brasileiro,
       ...; e CICLANO, brasileira, ...', identificados pelo padrão comum
       em instrumentos jurídicos brasileiros de citar cada pessoa como
       'NOME, brasileiro(a), ...'.
    Em todos os casos, o escopo fica limitado a partir de onde o documento
    começa a listar os procuradores (não pega nomes do Outorgante nem de
    quem o representa, que aparecem ANTES desse trecho)."""
    if not text:
        return []
    m_ini = _SECAO_OUTORGADOS_INICIO_RE.search(
        text
    ) or _SECAO_OUTORGADOS_INICIO_ALT_RE.search(text)
    if not m_ini:
        return []
    m_fim = _SECAO_OUTORGADOS_FIM_RE.search(text, m_ini.end())
    secao = (
        text[m_ini.end() : m_fim.start()]
        if m_fim
        else text[m_ini.end() : m_ini.end() + 3000]
    )

    nomes = [n.strip() for n in _NOME_LISTA_NUMERADA_RE.findall(secao)]
    if nomes:
        return nomes

    nomes = [" ".join(n.split()) for n in _NOME_BRASILEIRO_RE.findall(secao)]
    if nomes:
        return nomes

    m_nome = re.search(r"^\s*([A-ZÀ-Ý][^,\n\r]{2,80}),", secao, re.M)
    return [m_nome.group(1).strip()] if m_nome else []


def validate_procuracao(text: str, filename: str) -> DocumentValidation:
    """TAB415-004: 'Procuração outorgando poderes para assinar
    acordos/contratos e representar o fornecedor', insumo obrigatório no
    Cadastro/Atualização de Fornecedor Pessoa Jurídica. O orchestrator
    também usa a lista de outorgados pra não reprovar quando quem assina o
    Formulário é um procurador (em vez do sócio).
    """
    outorgante_match = OUTORGANTE_RE.search(text)
    outorgante_valor = outorgante_match.group(1).strip() if outorgante_match else ""

    outorgados = extrair_outorgados(text)
    if not outorgados:
        outorgado_match = OUTORGADO_RE.search(text)
        outorgado_valor = outorgado_match.group(1).strip() if outorgado_match else ""
        if _valor_parece_nome(outorgado_valor):
            outorgados = [outorgado_valor]
    return DocumentValidation(
        doc_type="PROCURACAO",
        filename=filename,
        valido=True,
        campos={
            "outorgante": outorgante_valor
            if _valor_parece_nome(outorgante_valor)
            else "",
            "outorgados": ";".join(outorgados),
        },
        pendencias=[],
    )
