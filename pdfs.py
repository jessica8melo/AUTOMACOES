#!/usr/bin/env python3
"""
Busca campos específicos dentro de um documento PDF, incluindo campos que
estão organizados em tabelas (com cabeçalhos quebrados em mais de uma linha)
e campos soltos no meio do texto.

Uso:
    python buscar_campos_pdf_v2.py caminho/do/arquivo.pdf
"""

import sys
import re
import unicodedata
from difflib import SequenceMatcher

import pdfplumber

try:
    import pytesseract
    OCR_DISPONIVEL = True
except ImportError:
    OCR_DISPONIVEL = False

# ---------------------------------------------------------------------------
# Campos que devem ser buscados no documento.
# Adicione, remova ou edite livremente.
# ---------------------------------------------------------------------------
CAMPOS_PROCURADOS = [
    "Valor total da solicitação",
    "Qtde",
    "Data do fornecimento",
    "Código",
    "Especificação do Bem",
    "Preço (s) unitário (s) (R$)",
    "DGCO nº",
    "Empresa",
    "OC Master nº"
]


# ---------------------------------------------------------------------------
# Utilidades de normalização e comparação "aproximada" de texto
# ---------------------------------------------------------------------------
def normalizar(texto: str) -> str:
    """Deixa o texto minúsculo, sem acentos, sem parênteses e sem pontuação."""
    if not texto:
        return ""
    texto = re.sub(r"\([^)]*\)", " ", texto)  # remove conteúdo entre parênteses
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def parecido(a: str, b: str, limite: float = 0.72) -> bool:
    """Compara dois textos de forma tolerante a pequenas variações de escrita."""
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return False
    if na == nb or na in nb or nb in na:
        return True
    return SequenceMatcher(None, na, nb).ratio() >= limite


def parece_valor(texto: str) -> bool:
    """Indica se um texto parece ser um VALOR (contém dígito) e não um rótulo."""
    return bool(re.search(r"\d", texto))


# ---------------------------------------------------------------------------
# Extração de texto e tabelas do PDF
# ---------------------------------------------------------------------------
def ocr_pagina(pagina, resolucao: int = 200) -> str:
    """
    IMPORTANTE: erros de OCR agora são avisados em stderr (em vez de
    engolidos em silêncio) para facilitar diagnosticar ambientes onde o
    Tesseract/pacote de idioma não está instalado.
    """
    if not OCR_DISPONIVEL:
        print(
            "[AVISO] pytesseract não está instalado neste ambiente — "
            "OCR desativado, páginas só-imagem (ex.: certificados de "
            "assinatura) não serão lidas.",
            file=sys.stderr,
        )
        return ""
    try:
        imagem = pagina.to_image(resolution=resolucao).original
    except Exception as erro:
        print(f"[AVISO] Falha ao converter página em imagem para OCR: {erro}", file=sys.stderr)
        return ""

    ultimo_erro = None
    for idioma in ("por", "eng", None):
        try:
            if idioma:
                return pytesseract.image_to_string(imagem, lang=idioma)
            return pytesseract.image_to_string(imagem)
        except Exception as erro:
            ultimo_erro = erro
            continue
    print(f"[AVISO] OCR falhou para todos os idiomas testados: {ultimo_erro}", file=sys.stderr)
    return ""


def _nao_eh_marca_dagua(obj) -> bool:
    """
    Usado em pagina.filter(...) para IGNORAR a marca d'água diagonal de
    rastreamento (ex.: e-mail de quem baixou o PDF + data/hora, repetida
    várias vezes na página, levemente rotacionada). Ela é impressa num
    cinza bem claro (quase branco) e seus caracteres acabam entrelaçados
    com o texto real na mesma região, embaralhando a extração (ex.:
    "Nota Técnica - 2022/0263" saindo como "m/-/1/2/9a/0/...").

    Devolve True para manter o objeto (texto normal) e False para
    descartá-lo (marca d'água). Só mexe em caracteres — outros objetos
    (linhas, retângulos, imagens) passam direto.
    """
    if obj.get("object_type") != "char":
        return True
    cor = obj.get("non_stroking_color")
    if not isinstance(cor, (tuple, list)) or len(cor) < 3:
        return True
    return not all(abs(c - 0.85882) < 0.02 for c in cor[:3])


def extrair_texto(caminho_pdf: str) -> str:
    partes = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            pagina_limpa = pagina.filter(_nao_eh_marca_dagua)
            texto_pagina = pagina_limpa.extract_text(layout=True) or ""
            if len(texto_pagina.strip()) < 200 and pagina.images:
                texto_ocr = ocr_pagina(pagina)
                if texto_ocr:
                    texto_pagina = f"{texto_pagina}\n{texto_ocr}"
            partes.append(texto_pagina)
    return "\n".join(partes)


def extrair_tabelas(caminho_pdf: str):
    tabelas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            tabelas.extend(pagina.extract_tables())
    return tabelas


def mapear_campos_em_tabela(tabela):
    linhas_info = []
    for linha in tabela:
        celulas = [(i, c.strip()) for i, c in enumerate(linha) if c and c.strip()]
        linhas_info.append(celulas)

    pares_rotulo_valor = []
    cabecalho_valor = {}

    i = 0
    total = len(linhas_info)
    while i < total:
        celulas = linhas_info[i]

        if not celulas:
            i += 1
            continue

        tem_valor = any(parece_valor(c) for _, c in celulas)

        if tem_valor and len(celulas) == 2 and not parece_valor(celulas[0][1]):
            pares_rotulo_valor.append((celulas[0][1], celulas[1][1]))
            i += 1
            continue

        if not tem_valor:
            bloco_cabecalho = [celulas]
            j = i + 1
            while j < total:
                proxima = linhas_info[j]
                if proxima and not any(parece_valor(c) for _, c in proxima):
                    bloco_cabecalho.append(proxima)
                    j += 1
                else:
                    break

            colunas = {}
            for celulas_linha in bloco_cabecalho:
                for col, texto in celulas_linha:
                    colunas[col] = (colunas.get(col, "") + " " + texto).strip()
            cabecalhos_em_ordem = [colunas[c] for c in sorted(colunas.keys())]

            while j < total:
                dados = linhas_info[j]
                if not dados:
                    break
                dados_tem_valor = any(parece_valor(c) for _, c in dados)
                if not dados_tem_valor:
                    break
                if len(dados) == 2 and not parece_valor(dados[0][1]):
                    break
                valores_em_ordem = [texto for _, texto in dados]
                for cabecalho, valor in zip(cabecalhos_em_ordem, valores_em_ordem):
                    cabecalho_valor.setdefault(cabecalho, valor)
                j += 1

            i = j
            continue

        i += 1

    return pares_rotulo_valor, cabecalho_valor


def buscar_em_tabelas(tabelas, campo: str):
    for tabela in tabelas:
        pares, cabecalho_valor = mapear_campos_em_tabela(tabela)

        for rotulo, valor in pares:
            if parecido(rotulo, campo):
                return valor

        for cabecalho, valor in cabecalho_valor.items():
            if parecido(cabecalho, campo):
                return valor

    return None


def buscar_em_texto(texto: str, campo: str):
    campo_escapado = re.escape(campo)

    padrao_mesma_linha = re.compile(rf"{campo_escapado}\s*[:\-]\s*(.+)", re.IGNORECASE)
    match = padrao_mesma_linha.search(texto)
    if match:
        valor = match.group(1).strip()
        if valor:
            return valor

    padrao_linha_seguinte = re.compile(rf"{campo_escapado}\s*\n\s*(.+)", re.IGNORECASE)
    match = padrao_linha_seguinte.search(texto)
    if match:
        valor = match.group(1).strip()
        if valor and (parece_valor(valor) or len(valor.split()) <= 6):
            return valor

    return None


PADROES_ESPECIAIS = {
    "Data do fornecimento": [
        r"firmad[ao]\s+e\s+assinad[ao].{0,80}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4})",
        r"assinad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
        r"firmad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
    ],
    "DGCO nº": [
        # \.? e :? extras: cobre "DGCO n.º", "DGCO nº:", "DGCO Nº :" etc.
        r"DGCO\s*n[ºo°]?\.?\s*:?\s*([\d./\-]+)",
    ],
    "OC Master nº": [
        r"OC\s*Master\s*n[ºo°]?\.?\s*:?\s*([\d./\-]+)",
        r"\bOC\s*n[ºo°]?\.?\s*:?\s*([\d./\-]+)",
    ],
    # A modalidade não aparece como rótulo "Tipo de Contratação: ..." — ela
    # é o próprio cabeçalho repetido em toda página (ex.: "LICITAÇÃO
    # ELETRÔNICA Nº 2022/04"). Captura só o nome da modalidade, sem o
    # número, mesmo que ele venha logo depois no texto.
    "Tipo de Contratação": [
        r"\b(LICITA[ÇC][ÃA]O\s+ELETR[ÔO]NICA|LICITA[ÇC][ÃA]O\s+PRESENCIAL"
        r"|PREG[ÃA]O\s+ELETR[ÔO]NICO|PREG[ÃA]O\s+PRESENCIAL"
        r"|DISPENSA\s+DE\s+LICITA[ÇC][ÃA]O|INEXIGIBILIDADE\s+DE\s+LICITA[ÇC][ÃA]O"
        r"|CONCORR[ÊE]NCIA|TOMADA\s+DE\s+PRE[ÇC]OS|CONVITE"
        r"|CONTRATA[ÇC][ÃA]O\s+DIRETA|ATA\s+DE\s+REGISTRO\s+DE\s+PRE[ÇC]OS)\b",
    ],
    # O contrato sempre cita 2 CNPJs (CONTRATANTE e CONTRATADA), na mesma
    # frase de qualificação das partes: "... inscrita no CNPJ nº
    # XX.XXX.XXX/XXXX-XX, denominada CONTRATADA". Interessa sempre o da
    # CONTRATADA (o fornecedor), nunca o da CONTRATANTE (BB Tecnologia).
    "CNPJ": [
        r"CNPJ\s*n[ºo°]?\.?\s*:?\s*([\d./\-]+)\s*,?\s*denominad[ao]\s+(?:a\s+)?CONTRATADA\b",
    ],
    # Aparece como cabeçalho no topo do documento: "Nota Técnica - 2022/0263"
    "Número da Nota Técnica": [
        r"Nota\s+T[ée]cnica\s*[-–]\s*(\d{4}\s*/\s*\d+)",
    ],
}


def buscar_com_padroes_especiais(texto: str, campo: str):
    """
    Antes: exigia que `campo` fosse EXATAMENTE igual (char a char) a uma
    chave de PADROES_ESPECIAIS — se o texto em CAMPOS_PROCURADOS estivesse
    escrito de forma levemente diferente (espaço a mais, "OC nº" em vez de
    "OC Master nº", etc.), a busca falhava silenciosamente.
    Agora usa a mesma comparação "aproximada" (parecido/normalizar) já usada
    no resto do script, então funciona mesmo com pequenas variações de texto.
    """
    for chave, padroes in PADROES_ESPECIAIS.items():
        if not parecido(chave, campo):
            continue
        for padrao in padroes:
            match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
            if match:
                return " ".join(match.group(1).split())
    return None


INDICADORES_ASSINATURA = [
    ("Aprovve", r"assinado\s+eletronicamente\s+atrav[eé]s\s+do\s+sistema\s+aprovve"),
    ("D4Sign", r"d4sign|certificado\s+pela\s+d4sign|padr[aã]o\s+icp[\s-]*brasil"),
    ("GOV.BR / ICP-Brasil", r"assinado\s+de\s+forma\s+eletr[oô]nica|verifique\s+em\s+https?://verificador\.iti\.gov\.br|assinatura\s+qualificada"),
    ("DocuSign", r"docusign"),
    ("Clicksign", r"clicksign"),
    ("Genérico", r"assinado\s+digitalmente|documento\s+assinado\s+eletronicamente"),
]

PADRAO_SIGNATARIO = re.compile(
    r"([A-ZÀ-Ú][A-ZÀ-Ú0-9\s/]+?)\s-\s(.+?)\s-\s(\d{2}/\d{2}/\d{4})\s[–-]\s(\d{2}:\d{2})"
)

PADRAO_SIGNATARIO_D4SIGN = re.compile(
    r"([A-ZÀ-Ú][A-ZÀ-Ú\s]{3,60}?)\s+"
    r"(Assinou como parte|Assinou como testemunha|Aprovou|Acusou recebimento|Rejeitou)\b"
    r".*?DATE[_ ]?ATOM:\s*(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2}):\d{2}",
    re.DOTALL,
)

PADRAO_NUMERO_PROCESSO = re.compile(
    r"sob\s+o\s+n[uú]mero\s+([\w./\-]+)", re.IGNORECASE
)

# A D4Sign não usa a frase "sob o número X"; ela identifica o envelope pelo
# "Código do documento <uuid>" impresso no certificado. Usado como
# alternativa quando PADRAO_NUMERO_PROCESSO não encontra nada.
PADRAO_CODIGO_D4SIGN = re.compile(
    r"[Cc][oó]digo\s+do\s+documento\s+([0-9a-fA-F\-]{20,40})"
)


def verificar_assinatura(texto: str) -> dict:
    resultado = {
        "assinado": False,
        "sistema": None,
        "numero_processo": None,
        "signatarios": [],
    }

    for nome_sistema, padrao in INDICADORES_ASSINATURA:
        if re.search(padrao, texto, re.IGNORECASE):
            resultado["assinado"] = True
            resultado["sistema"] = nome_sistema
            break

    match_numero = PADRAO_NUMERO_PROCESSO.search(texto)
    if match_numero:
        resultado["numero_processo"] = match_numero.group(1)
    else:
        match_codigo = PADRAO_CODIGO_D4SIGN.search(texto)
        if match_codigo:
            resultado["numero_processo"] = match_codigo.group(1)

    # Formato "NOME - CARGO - DD/MM/AAAA – HH:MM" (Aprovve e afins)
    for nome, cargo, data, hora in PADRAO_SIGNATARIO.findall(texto):
        resultado["signatarios"].append({
            "nome": " ".join(nome.split()),
            "cargo": " ".join(cargo.split()),
            "data": data,
            "hora": hora,
        })

    # Formato do bloco "Eventos do documento" da D4Sign (era definido mas
    # nunca chamado — por isso a assinatura era detectada como "assinado",
    # mas a lista de signatários ficava sempre vazia nesse tipo de PDF).
    ja_vistos = {(s["nome"], s["data"], s["hora"]) for s in resultado["signatarios"]}
    for nome, acao, ano, mes, dia, hora in PADRAO_SIGNATARIO_D4SIGN.findall(texto):
        data = f"{dia}/{mes}/{ano}"
        chave = (" ".join(nome.split()), data, hora)
        if chave in ja_vistos:
            continue
        ja_vistos.add(chave)
        resultado["signatarios"].append({
            "nome": " ".join(nome.split()),
            "cargo": acao,  # aqui "cargo" guarda a ação: Assinou / Aprovou / etc.
            "data": data,
            "hora": hora,
        })

    if resultado["signatarios"] and not resultado["assinado"]:
        resultado["assinado"] = True

    return resultado


def processar_pdf(caminho_pdf: str, campos_procurados: list = None) -> dict:
    """
    `campos_procurados` permite passar uma lista de campos específica
    (ex.: a combinação fluxo+documento decidida em main.py/doc_types.py)
    em vez da lista fixa CAMPOS_PROCURADOS. Se omitido, usa a lista fixa
    (mantém o comportamento antigo para quem chama/roda este arquivo
    isoladamente).
    """
    if campos_procurados is None:
        campos_procurados = CAMPOS_PROCURADOS

    texto = extrair_texto(caminho_pdf)
    tabelas = extrair_tabelas(caminho_pdf)

    resultado = {}
    for campo in campos_procurados:
        valor = buscar_em_tabelas(tabelas, campo)
        if not valor:
            valor = buscar_em_texto(texto, campo)
        if not valor:
            valor = buscar_com_padroes_especiais(texto, campo)

        resultado[campo] = " ".join(valor.split()) if valor else None

    resultado["_assinatura"] = verificar_assinatura(texto)

    return resultado


def imprimir_resultado(caminho_pdf: str, resultado: dict) -> None:
    print(f"Documento analisado: {caminho_pdf}\n")
    for campo, valor in resultado.items():
        if campo == "_assinatura":
            continue
        if valor:
            print(f"[SUCESSO] Campo '{campo}' encontrado. Valor: {valor}")
        else:
            print(f"[ERRO] Não foi possível encontrar o campo '{campo}' no documento.")

    assinatura = resultado.get("_assinatura", {})
    print()
    if assinatura.get("assinado"):
        sistema = assinatura.get("sistema") or "não identificado"
        print(f"[ASSINATURA] Documento assinado. Sistema: {sistema}")
        if assinatura.get("numero_processo"):
            print(f"[ASSINATURA] Nº do processo: {assinatura['numero_processo']}")
        for s in assinatura.get("signatarios", []):
            print(f"[ASSINATURA]   - {s['nome']} ({s['cargo']}) em {s['data']} {s['hora']}")
    else:
        print("[ASSINATURA] Nenhum indício de assinatura eletrônica encontrado no documento.")


def main():
    if len(sys.argv) < 2:
        print("Uso: python pdfs.py caminho/do/arquivo.pdf")
        sys.exit(1)

    caminho_pdf = sys.argv[1]

    try:
        resultado = processar_pdf(caminho_pdf)
    except Exception as erro:
        print(f"[ERRO] Não foi possível abrir/ler o arquivo '{caminho_pdf}': {erro}")
        sys.exit(1)

    imprimir_resultado(caminho_pdf, resultado)


if __name__ == "__main__":
    main()