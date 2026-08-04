#!/usr/bin/env python3
"""
Busca campos específicos dentro de um documento PDF, incluindo campos que
estão organizados em tabelas (com cabeçalhos quebrados em mais de uma linha)
e campos soltos no meio do texto.
"""

import sys
import re
import unicodedata
import logging
from difflib import SequenceMatcher

import pdfplumber
import pypdf

logging.getLogger("pdfminer").setLevel(logging.ERROR)

try:
    import os
    import platform
    import pytesseract
    OCR_DISPONIVEL = True

    if platform.system() == "Windows":
        caminhos_padrao = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        local_appdata = os.environ.get("LOCALAPPDATA")
        if local_appdata:
            caminhos_padrao.append(
                os.path.join(local_appdata, "Tesseract-OCR", "tesseract.exe")
            )
        for caminho in caminhos_padrao:
            if os.path.isfile(caminho):
                pytesseract.pytesseract.tesseract_cmd = caminho
                break
except ImportError:
    OCR_DISPONIVEL = False


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
    """Compara dois textos de forma tolerante a pequenas variações de escrita.

    IMPORTANTE (1): a checagem "um texto está contido no outro" só vale para
    textos com pelo menos 4 caracteres normalizados. Sem esse piso, uma
    célula de tabela de 1-2 letras (lixo comum de PDF, ex.: 'o', 'no', 'c',
    restos de "...de sobre-aviso") acaba "contida" em QUALQUER campo mais
    longo (ex.: "Número da Nota Técnica" contém a letra 'o'), fazendo o
    campo casar com a célula errada.

    IMPORTANTE (2): mesmo com 4+ caracteres, uma palavra genérica de uma
    tabela qualquer (ex.: cabeçalho "ITEM" de uma tabela de multas, numerando
    infrações 1-5) NÃO deve "vencer" só por estar contida em um campo
    composto e bem mais específico como "Código do Item" ou "Descrição do
    Item" — ela poderia estar contida em qualquer campo que mencione "item"
    em outro contexto qualquer. Por isso a inclusão só conta quando o texto
    curto tem 2+ palavras (é um rótulo, não uma palavra solta) OU cobre pelo
    menos metade do comprimento do texto longo (não é um fragmento pequeno
    perdido dentro de um rótulo bem maior).
    """
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    curto, longo = (na, nb) if len(na) <= len(nb) else (nb, na)
    if len(curto) >= 4 and curto in longo:
        if len(curto.split()) >= 2 or len(curto) / len(longo) >= 0.5:
            return True
    
    return SequenceMatcher(None, na, nb).ratio() >= max(limite, 0.85)


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
    rastreamento.

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

AREA_MINIMA_IMAGEM_CARIMBO = 5000

def _pagina_tem_imagem_grande(pagina, area_minima: float = AREA_MINIMA_IMAGEM_CARIMBO) -> bool:
    """
    True se a página tiver alguma imagem grande o bastante para ser um
    carimbo de assinatura (ex.: aquele selo com "Assinado de forma
    digital por..."). Usado para decidir se vale rodar OCR mesmo em
    páginas que já têm bastante texto normal — o carimbo pode conviver
    com texto real na mesma página (ex.: um Projeto Básico com campos
    normais + assinatura ao final), e nesse caso o texto do carimbo em
    si nunca aparece na extração normal de texto, só na imagem.
    """
    for imagem in pagina.images:
        largura = imagem.get("width", 0) or 0
        altura = imagem.get("height", 0) or 0
        if largura * altura >= area_minima:
            return True
    return False


def extrair_texto(caminho_pdf: str) -> str:
    partes = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            pagina_limpa = pagina.filter(_nao_eh_marca_dagua)
            texto_pagina = pagina_limpa.extract_text(layout=True) or ""
            precisa_ocr = pagina.images and (
                len(texto_pagina.strip()) < 200 or _pagina_tem_imagem_grande(pagina)
            )
            if precisa_ocr:
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
    """
    Reúne TODOS os valores encontrados no documento
    inteiro para esse campo. Se todos forem iguais (o mesmo valor
    aparecendo em mais de uma tabela, ou uma única tabela batendo), o
    valor é devolvido normalmente — não há ambiguidade real. Mas se
    houver valores DIFERENTES vindos de tabelas diferentes, isso é sinal
    de que o campo é genérico demais para esse documento (não tem *um*
    valor de "Quantidade" que faça sentido no nível do contrato) — nesse
    caso, é mais honesto devolver None (o chamador reporta "não
    encontrado") do que inventar uma resposta a partir da tabela que por
    acaso apareceu primeiro.
    """
    valores_encontrados = []

    for tabela in tabelas:
        pares, cabecalho_valor = mapear_campos_em_tabela(tabela)

        for rotulo, valor in pares:
            if parecido(rotulo, campo):
                valores_encontrados.append(valor)

        for cabecalho, valor in cabecalho_valor.items():
            if parecido(cabecalho, campo):
                valores_encontrados.append(valor)

    valores_unicos = list(dict.fromkeys(v for v in valores_encontrados if v))
    if len(valores_unicos) == 1:
        return valores_unicos[0]
    return None


def buscar_em_texto(texto: str, campo: str):
    campo_escapado = re.escape(campo)
    padrao_mesma_linha = re.compile(
        rf"\b{campo_escapado}\b\s*[:\-]\s*(.+?)"
        rf"(?=\n\s*\n|\n\s*\d+(?:\.\d+)*[\.\)]|\n\s*[A-ZÀ-Ý][^\n:]{{0,80}}:|\Z)",
        re.IGNORECASE | re.DOTALL,
    )
    match = padrao_mesma_linha.search(texto)
    if match:
        valor = match.group(1).strip()
        if valor:
            return valor

    padrao_linha_seguinte = re.compile(rf"\b{campo_escapado}\b\s*\n\s*(.+)", re.IGNORECASE)
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
    "Data do contrato": [
        r"celebrad[ao]\s+pelas\s+partes\s+em\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4})",
        r"celebrad[ao].{0,80}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
        r"firmad[ao]\s+e\s+assinad[ao].{0,80}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4})",
        r"assinad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
        r"firmad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
    ],
    "DGCO nº": [
        r"DGCO\s*(?:n[ºo°]?\.?)?\s*:?\s*([\d./\-]+)",
    ],
    "OC Master nº": [
        r"OC\s*Master\s*(?:n[ºo°]?\.?)?\s*:?\s*([\d./\-]+)",
        r"\bOC\s*(?:n[ºo°]?\.?)?\s*:?\s*([\d./\-]+)",
    ],
    "Tipo de Contratação": [
        r"\b(LICITA[ÇC][ÃA]O\s+ELETR[ÔO]NICA|LICITA[ÇC][ÃA]O\s+PRESENCIAL"
        r"|PREG[ÃA]O\s+ELETR[ÔO]NICO|PREG[ÃA]O\s+PRESENCIAL"
        r"|DISPENSA\s+DE\s+LICITA[ÇC][ÃA]O|INEXIGIBILIDADE\s+DE\s+LICITA[ÇC][ÃA]O"
        r"|CONCORR[ÊE]NCIA|TOMADA\s+DE\s+PRE[ÇC]OS|CONVITE"
        r"|CONTRATA[ÇC][ÃA]O\s+DIRETA|ATA\s+DE\s+REGISTRO\s+DE\s+PRE[ÇC]OS)\b",
    ],
    "CNPJ": [
        r"CNPJ\s*(?:n[ºo°]?\.?)?\s*:?\s*([\d./\-]+)\s*,?\s*(?:doravante\s+)?denominad[ao]\s+(?:a\s+)?CONTRATADA\b",
    ],
    "Contratada": [
        r"([A-ZÀ-Ü][A-Za-zÀ-ÿ0-9°º\.,'&/\-\s]{2,120}?),?\s*(?:pessoa\s+jur[íi]dica[^,]{0,80},)?\s*"
        r"(?:inscrita|estabelecida)[^,]{0,80}?CNPJ\s*(?:n[ºo°]?\.?)?\s*:?\s*[\d./\-]+\s*,?\s*"
        r"(?:doravante\s+)?denominad[ao]\s+(?:a\s+)?CONTRATADA\b",
    ],
    "Número da Nota Técnica": [
        r"Nota\s+T[ée]cnica\s*[-–]\s*(\d{4}\s*/\s*\d+)",
    ],
    "Condições de Pagamento": [
        r"pagament\w*.{0,300}?\b(?:em|at[ée])\s+(?:at[ée]\s+)?(?:o\s+)?(\d+\s*[ºo°]?\s*(?:\([^)]*\))?\s*dias?(?:\s+corridos)?(?:\s+do\s+m[êe]s\s+subsequente)?)",
        r"(\d+\s*(?:\([^)]*\))?\s*dias\s+corridos)(?=.{0,150}?(?:emiss[ãa]o\s+da\s+nota\s+fiscal|conclus[ãa]o\s+d[eas]))",
    ],
    "Valor total": [
        r"valor\s+total\s+do\s+contrato\s+passar[áa]\s+de\s+R\$\s*[\d.,]+\s+para\s+(?:o\s+valor\s+total\s+de\s+)?R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
        r"(?:no\s+)?valor\s+total\s+(?:do\s+contrato\s+)?(?:estimado\s+)?(?:(?:[ée]|ser[áa])\s+)?de\s+(?:at[ée]\s+)?R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
    ],
}

# Campos cuja busca via padrão especial (regex) deve ter PRIORIDADE sobre a
# busca genérica de tabela/"rótulo: valor". Normalmente os padrões especiais
# só entram como último recurso (quando tabela e texto genérico falham),
# mas para estes campos a busca genérica encontraria algo tecnicamente
# "correto" — um trecho de texto que realmente segue o rótulo — só que não
# é o valor que interessa (ex.: "Condições de Pagamento" tem um parágrafo
# de regras logo após o rótulo, mas o que se quer é o prazo em dias
# corridos, que pode estar em outro parágrafo do documento).
CAMPOS_PRIORIDADE_ESPECIAL = {
    "Condições de Pagamento",
    "Valor total",
    "Contratada",
}


def _extrair_contratada_por_coluna(texto: str):
    """
    No modelo de contrato da BBTS, a qualificação das partes vem em duas
    caixas lado a lado — CONTRATANTE (esquerda) e CONTRATADA (direita) —,
    cada uma com "RAZÃO SOCIAL: ... / NOME FANTASIA: ... / CNPJ: ... /
    ENDEREÇO: ...". Como o texto é extraído com extract_text(layout=True)
    (preserva a posição horizontal de cada palavra na página), as duas
    colunas ficam intercaladas na MESMA linha de texto — ex.:
    "RAZÃO SOCIAL: BB TECNOLOGIA E      RAZÃO SOCIAL: SANT'COSTA LIMPEZA".
    Uma regex simples de "ache a 2ª ocorrência de RAZÃO SOCIAL" falha
    porque as duas ocorrências nem sempre ficam em linhas diferentes, e a
    caixa da CONTRATADA costuma ter uma linha a menos (não tem "NOME
    FANTASIA"), desalinhando o restante das linhas entre as duas colunas.

    A abordagem aqui usa a POSIÇÃO horizontal (coluna de caractere) onde a
    palavra "CONTRATADA" aparece no cabeçalho "CONTRATANTE   CONTRATADA"
    para saber, em cada linha seguinte, a partir de qual coluna cortar o
    texto — isolando só o conteúdo da caixa da direita (CONTRATADA) antes
    de procurar "RAZÃO SOCIAL:" nele.
    """
    linhas = texto.split("\n")

    coluna_contratada = None
    indice_cabecalho = None
    for i, linha in enumerate(linhas):
        match_cabecalho = re.search(r"CONTRATANTE\s{2,}CONTRATADA\b", linha)
        if match_cabecalho:
            coluna_contratada = linha.index("CONTRATADA", match_cabecalho.start())
            indice_cabecalho = i
            break

    if coluna_contratada is None:
        return None

    corte = max(coluna_contratada - 8, 0)

    linhas_direita = []
    for linha in linhas[indice_cabecalho + 1: indice_cabecalho + 15]:
        if not linha.strip():
            if linhas_direita:
                break
            continue
        linhas_direita.append(linha[corte:])

    bloco_contratada = "\n".join(linhas_direita)

    match = re.search(
        r"RAZ[ÃA]O\s+SOCIAL\s*:\s*(.+?)\s*\n\s*(?:CNPJ|NOME\s+FANTASIA|ENDERE[ÇC]O)\b",
        bloco_contratada,
        re.IGNORECASE | re.DOTALL,
    )
    if match:
        valor = " ".join(match.group(1).split())
        if valor:
            return valor

    return None


def buscar_com_padroes_especiais(texto: str, campo: str):
    """
    Usa a mesma comparação "aproximada" (parecido/normalizar) já usada
    no resto do script, então funciona mesmo com pequenas variações de texto.
    """
    if parecido("Contratada", campo):
        valor_coluna = _extrair_contratada_por_coluna(texto)
        if valor_coluna:
            return valor_coluna

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
    ("BBTS (interno)", r"assinado\s+eletronicamente\s+por\s*:"),
    ("Adobe Acrobat (certificado digital)", r"assinado\s+de\s+forma\s+digital\s+por"),
    ("Genérico", r"assinado\s+digitalmente|documento\s+assinado\s+eletronicamente"),
]

PADRAO_SIGNATARIO = re.compile(
    r"([A-ZÀ-Ú][A-ZÀ-Ú0-9\s/]+?)\s-\s(.+?)\s-\s(\d{2}/\d{2}/\d{4})\s[–-]\s(\d{2}:\d{2})"
)

PADRAO_SIGNATARIO_BBTS = re.compile(
    r"Assinado\s+eletronicamente\s+por:\s*"
    r"([A-ZÀ-Ú][A-ZÀ-Ú0-9\s/]+?),\s*em\s+(\d{2}/\d{2}/\d{4})\s+(?:[àa]s\s+)?(\d{2}:\d{2})"
    r"\s*\n\s*(.+)",
    re.IGNORECASE,
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

PADRAO_CODIGO_D4SIGN = re.compile(
    r"[Cc][oó]digo\s+do\s+documento\s+([0-9a-fA-F\-]{20,40})"
)

PADRAO_SIGNATARIO_ADOBE = re.compile(
    r"Assinado\s+de\s+forma\s+digital\s+por\s+"
    r"(.+?)\s*\(([\w.\-]+)\)"
    r".{0,40}?Dados:\s*(\d{4})\.(\d{2})\.(\d{2})\s+(\d{2}:\d{2}:\d{2})",
    re.IGNORECASE | re.DOTALL,
)


PADRAO_DATA_ASSINATURA_PDF = re.compile(
    r"D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})"
)


def verificar_assinaturas_acroform(caminho_pdf: str) -> dict:
    """
    Lê as assinaturas digitais "de verdade" do PDF: os campos /Sig do
    AcroForm, que guardam a assinatura criptográfica (certificado PKI —
    ICP-Brasil, e-CPF/e-CNPJ via Adobe, por ex.) independente de haver ou
    não um carimbo visual desenhado na página.

    Isso cobre um caso que nem o texto normal nem o OCR de página
    conseguem pegar: o carimbo "Assinado de forma digital por NOME
    (usuario) / Dados: AAAA.MM.DD HH:MM:SS ±HH'MM'" que o Adobe Reader
    mostra é uma STRING PRÓPRIA DO ADOBE, montada a partir do dicionário
    /V de cada campo /Sig (chaves /Name e /M) — ela não fica "desenhada"
    nos pixels da página nem no conteúdo/anotações do PDF em si, então
    nenhum extrator de texto ou OCR de página consegue enxergá-la. Esse
    dicionário, porém, sempre existe no PDF (é o que garante a validade
    jurídica da assinatura), então lê-lo diretamente é mais confiável do
    que tentar reconhecer o carimbo visual.
    """
    resultado = {"assinado": False, "sistema": None, "signatarios": []}
    try:
        leitor = pypdf.PdfReader(caminho_pdf)
    except Exception as erro:
        print(f"[AVISO] Falha ao abrir PDF para checar assinaturas do AcroForm: {erro}", file=sys.stderr)
        return resultado

    acroform = leitor.trailer.get("/Root", {}).get("/AcroForm")
    campos = acroform.get("/Fields") if acroform else None
    if not campos:
        return resultado

    for campo in campos:
        obj = campo.get_object()
        if obj.get("/FT") != "/Sig":
            continue
        valor = obj.get("/V")
        if not valor:
            continue  # campo de assinatura existe no formulário mas ainda não foi assinado
        v = valor.get_object()
        nome_bruto = v.get("/Name")
        if not nome_bruto:
            continue

        resultado["assinado"] = True
        if not resultado["sistema"]:
            filtro = str(v.get("/Filter", ""))
            resultado["sistema"] = (
                "Adobe/PKI (certificado digital)" if "Adobe" in filtro else (filtro.strip("/") or "PKI")
            )

        # O /Name normalmente vem como "NOME (usuario)" — separa os dois.
        match_nome = re.match(r"(.+?)\s*\(([\w.\-]+)\)\s*$", str(nome_bruto))
        nome = match_nome.group(1).strip() if match_nome else str(nome_bruto)
        usuario = match_nome.group(2) if match_nome else ""

        data_str, hora_str = "", ""
        m_data = v.get("/M")
        if m_data:
            match_data = PADRAO_DATA_ASSINATURA_PDF.match(str(m_data))
            if match_data:
                ano, mes, dia, h, mi, _s = match_data.groups()
                data_str = f"{dia}/{mes}/{ano}"
                hora_str = f"{h}:{mi}"

        resultado["signatarios"].append({
            "nome": nome,
            "cargo": usuario,  # aqui "cargo" guarda o usuário/login do certificado
            "data": data_str,
            "hora": hora_str,
        })

    return resultado


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

    ja_vistos_bbts = {(s["nome"], s["data"], s["hora"]) for s in resultado["signatarios"]}
    for nome, data, hora, cargo in PADRAO_SIGNATARIO_BBTS.findall(texto):
        nome_normalizado = " ".join(nome.split())
        chave = (nome_normalizado, data, hora)
        if chave in ja_vistos_bbts:
            continue
        ja_vistos_bbts.add(chave)
        resultado["signatarios"].append({
            "nome": nome_normalizado,
            "cargo": " ".join(cargo.split()),
            "data": data,
            "hora": hora,
        })

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

    ja_vistos_adobe = {(s["nome"], s["data"], s["hora"]) for s in resultado["signatarios"]}
    for nome, usuario, ano, mes, dia, hora in PADRAO_SIGNATARIO_ADOBE.findall(texto):
        nome_normalizado = " ".join(nome.split())
        data = f"{dia}/{mes}/{ano}"
        chave = (nome_normalizado, data, hora)
        if chave in ja_vistos_adobe:
            continue
        ja_vistos_adobe.add(chave)
        resultado["signatarios"].append({
            "nome": nome_normalizado,
            "cargo": usuario,  # aqui "cargo" guarda o usuário/login do certificado
            "data": data,
            "hora": hora,
        })

    if resultado["signatarios"] and not resultado["assinado"]:
        resultado["assinado"] = True

    return resultado


def processar_pdf(caminho_pdf: str, campos_procurados: list = None) -> dict:
    """
    `campos_procurados` permite passar uma lista de campos específica
    (ex.: a combinação fluxo+documento decidida em arquivo.py/doc_types.py).
    """
    texto = extrair_texto(caminho_pdf)
    tabelas = extrair_tabelas(caminho_pdf)

    resultado = {}
    for campo in campos_procurados:
        prioridade_especial = any(
            parecido(campo, c) for c in CAMPOS_PRIORIDADE_ESPECIAL
        )

        valor = None
        if prioridade_especial:
            valor = buscar_com_padroes_especiais(texto, campo)

        if not valor:
            valor = buscar_em_tabelas(tabelas, campo)
        if not valor:
            valor = buscar_em_texto(texto, campo)
        if not valor:
            valor = buscar_com_padroes_especiais(texto, campo)

        resultado[campo] = " ".join(valor.split()) if valor else None

    assinatura_acroform = verificar_assinaturas_acroform(caminho_pdf)
    assinatura_texto = verificar_assinatura(texto)

    signatarios_combinados = list(assinatura_acroform["signatarios"])
    ja_vistos = {(s["nome"], s["data"], s["hora"]) for s in signatarios_combinados}
    for s in assinatura_texto["signatarios"]:
        chave = (s["nome"], s["data"], s["hora"])
        if chave in ja_vistos:
            continue
        ja_vistos.add(chave)
        signatarios_combinados.append(s)

    resultado["_assinatura"] = {
        "assinado": assinatura_acroform["assinado"] or assinatura_texto["assinado"],
        "sistema": assinatura_acroform["sistema"] or assinatura_texto["sistema"],
        "numero_processo": assinatura_texto["numero_processo"],
        "signatarios": signatarios_combinados,
    }

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
