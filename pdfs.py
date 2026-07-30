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
import logging
from difflib import SequenceMatcher

import pdfplumber
import pypdf

# O pdfminer (usado por baixo dos panos pelo pdfplumber) solta avisos no log
# quando encontra comandos de cor "estranhos" dentro do PDF — o caso mais
# comum é um preenchimento com PADRÃO (Pattern), tipo "/P1 scn", que ele
# tenta (sem sucesso) interpretar como um número de nível de cinza: "Cannot
# set non-stroke color: '/P1' is an invalid float value". Isso é só um
# aviso — o pdfminer ignora aquele comando específico e segue extraindo o
# resto do texto/tabelas normalmente — mas polui a saída do checklist sem
# agregar nada útil. Por isso o logger do pdfminer fica restrito a ERROR
# (avisos/warnings deixam de ser exibidos; erros de verdade continuam).
logging.getLogger("pdfminer").setLevel(logging.ERROR)

try:
    import os
    import platform
    import pytesseract
    OCR_DISPONIVEL = True

    # No Windows o Tesseract não entra no PATH sozinho (diferente do
    # macOS/Linux via brew/apt). Se o binário configurado não for
    # encontrado, tenta os locais padrão do instalador oficial antes de
    # desistir — assim não é preciso mexer manualmente na variável PATH.
    if platform.system() == "Windows":
        caminhos_padrao = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        # Instalação sem privilégio de administrador (comum quando o
        # instalador é rodado sem permissão para gravar em Program Files)
        # cai em AppData\Local do usuário logado.
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


# Área mínima (largura x altura, em pontos PDF) para uma imagem ser
# considerada "grande o bastante para ser um carimbo de assinatura".
# Escolhida para pegar carimbos típicos (ex.: 373x288 = ~107k) e ignorar
# logos/cabeçalhos finos de página (ex.: 146x17 = ~2.5k).
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
            # Antes, o OCR só rodava quando a página inteira tinha pouco
            # texto (< 200 caracteres), assumindo que era uma página
            # "só imagem". Isso deixava passar carimbos de assinatura
            # (ex.: selo do Adobe Acrobat/certificado digital) em páginas
            # que já têm bastante texto normal (valores, cláusulas etc.) —
            # o carimbo é só mais uma imagem na página, e seu texto nunca
            # é lido. Agora também rodamos OCR quando há uma imagem grande
            # o suficiente para ser um carimbo, mesmo com texto normal
            # abundante na página.
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

    # IMPORTANTE: o valor de um campo às vezes quebra de linha no meio da
    # extração do PDF (ex.: "...Suporte Técnico: não se\n      aplica."),
    # porque pdfplumber preserva o layout visual da página. Sem DOTALL, "."
    # não cruza "\n" e o valor saía cortado ("não se", perdendo "aplica.").
    # Por isso agora o valor pode continuar por mais linhas, mas para no
    # primeiro sinal claro de que acabou: uma linha em branco, ou o início
    # do próximo item numerado (ex.: "9.", "10.1", "2.3.2.1.4.2"), ou o fim
    # do texto — o que vier primeiro.
    # \b nas duas pontas: sem isso, um campo curto como "UOR" também
    # "casava" como SUBSTRING dentro de qualquer palavra que contivesse
    # essas letras em sequência (ex.: "LÍQUOR", em "...VÍRUS HERPES 6
    # LÍQUOR - IGG/IGM..." — nada a ver com o campo UOR do documento) e
    # aí o valor capturado era o que viesse depois do "-", uma lista de
    # exames sem relação nenhuma com o campo procurado.
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
    # A "Data do contrato" nunca aparece como rótulo solto ("Data do
    # contrato: ..."). Em Contratos, costuma vir na fórmula de fecho
    # ("...e assinam o presente contrato em DD de mês de AAAA..."). Em
    # Aditivos, o contrato original já está em vigor — a data que importa
    # é a de celebração do PRIMITIVO, citada na cláusula de Ratificação
    # ("...contrato de prestação de serviços... celebrado pelas partes em
    # 17 de março de 2022 e seus respectivos aditivos..."). Por isso os
    # padrões abaixo cobrem tanto "celebrado em" (Aditivos) quanto
    # "firmado/assinado em" (Contratos), nessa ordem de prioridade.
    "Data do contrato": [
        r"celebrad[ao]\s+pelas\s+partes\s+em\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4})",
        r"celebrad[ao].{0,80}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
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
        r"CNPJ\s*n[ºo°]?\.?\s*:?\s*([\d./\-]+)\s*,?\s*(?:doravante\s+)?denominad[ao]\s+(?:a\s+)?CONTRATADA\b",
    ],
    # Nome/Razão Social do FORNECEDOR (a CONTRATADA), não o texto qualquer
    # que segue a palavra "CONTRATADA" no meio do contrato (ex.: cláusula
    # de pagamento "A CONTRATANTE pagará à CONTRATADA:" — isso NÃO é o
    # nome da empresa, mas a busca genérica por rótulo "Contratada:"
    # casava justamente com essa cláusula, por ser a primeira ocorrência
    # de "CONTRATADA" seguida de dois-pontos no texto).
    #
    # Fórmula clássica de qualificação: "<empresa>, ..., inscrita no CNPJ
    # nº XX.XXX.XXX/XXXX-XX, ... denominada CONTRATADA". Cobre contratos
    # em texto corrido (fora do modelo de caixas lado a lado da BBTS, que
    # é tratado à parte por _extrair_contratada_por_coluna, chamada antes
    # deste padrão em buscar_com_padroes_especiais).
    "Contratada": [
        r"([A-ZÀ-Ü][A-Za-zÀ-ÿ0-9°º\.,'&/\-\s]{2,120}?),?\s*(?:pessoa\s+jur[íi]dica[^,]{0,80},)?\s*"
        r"(?:inscrita|estabelecida)[^,]{0,80}?CNPJ\s*n[ºo°]?\.?\s*:?\s*[\d./\-]+\s*,?\s*"
        r"(?:doravante\s+)?denominad[ao]\s+(?:a\s+)?CONTRATADA\b",
    ],
    # Aparece como cabeçalho no topo do documento: "Nota Técnica - 2022/0263"
    "Número da Nota Técnica": [
        r"Nota\s+T[ée]cnica\s*[-–]\s*(\d{4}\s*/\s*\d+)",
    ],
    # O campo "Condições de Pagamento" não deve trazer o parágrafo inteiro
    # do item "10. Condições de Pagamento" (que normalmente só descreve o
    # rito/gatilho do pagamento, ex.: "ao final de cada etapa"). O que
    # interessa é o PRAZO em dias corridos citado em algum ponto do texto
    # próximo à palavra "pagamento" (ex.: "...será realizado em até 30
    # (trinta) dias corridos..." ou "...em 30 (trinta) dias corridos, a
    # contar da emissão da Nota Fiscal..."). Por isso este campo tem
    # prioridade especial (ver CAMPOS_PRIORIDADE_ESPECIAL) e busca
    # diretamente por esse trecho, em vez do rótulo "Condições de
    # Pagamento:" seguido do parágrafo genérico.
    "Condições de Pagamento": [
        # Cobre as variações mais comuns de como o prazo aparece perto da
        # palavra "pagamento": "em até 30 (trinta) dias corridos", "até o
        # 15º (décimo quinto) dia do mês subsequente", ou só "em 30 dias"
        # (sem "corridos"). A ideia (pedido do usuário): não trazer a
        # primeira frase do parágrafo de "Condições de Pagamento" — ir
        # direto no primeiro trecho que fala de PRAZO em dias.
        r"pagament\w*.{0,300}?\b(?:em|at[ée])\s+(?:at[ée]\s+)?(?:o\s+)?(\d+\s*[ºo°]?\s*(?:\([^)]*\))?\s*dias?(?:\s+corridos)?(?:\s+do\s+m[êe]s\s+subsequente)?)",
        # Fallback: "dias corridos" que aparece perto de "emissão da nota
        # fiscal"/"conclusão de", mesmo sem a palavra "pagamento" por perto.
        r"(\d+\s*(?:\([^)]*\))?\s*dias\s+corridos)(?=.{0,150}?(?:emiss[ãa]o\s+da\s+nota\s+fiscal|conclus[ãa]o\s+d[eas]))",
    ],
    # Em aditivos de repactuação, o valor total aparece na frase "O valor
    # total do contrato passará de R$ X para o valor total de R$ Y" — o
    # que interessa é sempre o valor NOVO (Y, depois do "para"), nunca o
    # valor antigo que vem antes dele na mesma frase. Por isso tem
    # prioridade especial (ver CAMPOS_PRIORIDADE_ESPECIAL): a busca
    # genérica por rótulo pegaria o primeiro "R$ ..." que aparece perto de
    # "valor total", que costuma ser justamente o valor antigo.
    # O segundo padrão cobre contratos sem repactuação, onde o valor total
    # só aparece uma vez ("no valor total de R$ X" / "valor total é de
    # R$ X").
    "Valor total": [
        r"valor\s+total\s+do\s+contrato\s+passar[áa]\s+de\s+R\$\s*[\d.,]+\s+para\s+(?:o\s+valor\s+total\s+de\s+)?R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
        r"(?:no\s+)?valor\s+total\s+(?:do\s+contrato\s+)?(?:[ée]|ser[áa])\s+de\s+R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
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
    # Sem isso, a busca genérica por rótulo "Contratada: ..." casa com a
    # primeira ocorrência de "CONTRATADA" seguida de ":" no texto — que
    # normalmente é a cláusula de pagamento ("A CONTRATANTE pagará à
    # CONTRATADA:"), não a qualificação das partes. Ver padrão especial
    # "Contratada" em PADROES_ESPECIAIS.
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

    # Pequena margem para a esquerda: a posição da coluna varia alguns
    # caracteres entre linhas (estimativa de largura de fonte do
    # pdfplumber), sem risco de invadir o texto da CONTRATANTE, que fica
    # bem mais à esquerda.
    corte = max(coluna_contratada - 8, 0)

    linhas_direita = []
    for linha in linhas[indice_cabecalho + 1: indice_cabecalho + 15]:
        if not linha.strip():
            # caixa termina na primeira linha em branco após o cabeçalho
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
    Antes: exigia que `campo` fosse EXATAMENTE igual (char a char) a uma
    chave de PADROES_ESPECIAIS — se o texto em CAMPOS_PROCURADOS estivesse
    escrito de forma levemente diferente (espaço a mais, "OC nº" em vez de
    "OC Master nº", etc.), a busca falhava silenciosamente.
    Agora usa a mesma comparação "aproximada" (parecido/normalizar) já usada
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
    # Bloco interno da BBTS (ex.: Notas Técnicas): "Assinado eletronicamente
    # por: NOME, em DD/MM/AAAA às HH:MM" seguido do cargo/área na(s) linha(s)
    # de baixo. Precisa vir ANTES do "Genérico" porque a frase é parecida
    # ("assinado eletronicamente"), mas na ordem "assinado eletronicamente
    # POR", não "documento assinado eletronicamente".
    ("BBTS (interno)", r"assinado\s+eletronicamente\s+por\s*:"),
    # Carimbo padrão do Adobe Acrobat para certificado digital (ICP-Brasil
    # via e-CPF/e-CNPJ, por ex.): "Assinado de forma digital por NOME
    # (usuario)\nDados: AAAA.MM.DD HH:MM:SS ±HH'MM'". Geralmente vem como
    # imagem/carimbo na página (não como texto normal do PDF), então só é
    # capturado quando a página passa pelo OCR — ver _pagina_tem_imagem_grande
    # em extrair_texto.
    ("Adobe Acrobat (certificado digital)", r"assinado\s+de\s+forma\s+digital\s+por"),
    ("Genérico", r"assinado\s+digitalmente|documento\s+assinado\s+eletronicamente"),
]

PADRAO_SIGNATARIO = re.compile(
    r"([A-ZÀ-Ú][A-ZÀ-Ú0-9\s/]+?)\s-\s(.+?)\s-\s(\d{2}/\d{2}/\d{4})\s[–-]\s(\d{2}:\d{2})"
)

# Bloco interno da BBTS: "Assinado eletronicamente por: NOME, em DD/MM/AAAA
# às HH:MM", com o CARGO na linha logo abaixo (ex.: "GERENTE DE DIVISAO").
# A palavra "às" às vezes some (mesmo dentro do MESMO documento, ex.: os
# primeiros signatários de uma Nota Técnica têm "às" e os últimos não) —
# por isso ela é opcional aqui.
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

# A D4Sign não usa a frase "sob o número X"; ela identifica o envelope pelo
# "Código do documento <uuid>" impresso no certificado. Usado como
# alternativa quando PADRAO_NUMERO_PROCESSO não encontra nada.
PADRAO_CODIGO_D4SIGN = re.compile(
    r"[Cc][oó]digo\s+do\s+documento\s+([0-9a-fA-F\-]{20,40})"
)

# Carimbo do Adobe Acrobat: "Assinado de forma digital por NOME (usuario)
# Dados: AAAA.MM.DD HH:MM:SS ±HH'MM'". Como esse texto normalmente só
# aparece via OCR (o carimbo é uma imagem), o layout pode sair um pouco
# fora de ordem — por isso o casamento é tolerante a quebras de linha
# extras entre as partes (re.DOTALL) e não exige a data logo em seguida.
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

    # Formato interno da BBTS: "Assinado eletronicamente por: NOME, em
    # DD/MM/AAAA às HH:MM" com o cargo na linha seguinte (ex.: Notas
    # Técnicas). Era o único formato de assinatura sem suporte nenhum —
    # nem entrava no "assinado" (nenhum INDICADOR batia) nem tinha
    # signatário extraído.
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

    # Carimbo do Adobe Acrobat / certificado digital (ICP-Brasil): "Assinado
    # de forma digital por NOME (usuario) Dados: AAAA.MM.DD HH:MM:SS".
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

    # Duas fontes de verdade, combinadas: o AcroForm (/Sig) é a mais
    # confiável — é a assinatura criptográfica de fato, presente mesmo
    # quando não há nenhum carimbo visual na página (caso do Adobe
    # Reader). O texto/OCR pega sistemas que carimbam texto de verdade
    # na página (D4Sign, Aprovve, bloco interno da BBTS etc.), incluindo
    # casos sem AcroForm nenhum. Um documento é considerado assinado se
    # qualquer uma das duas fontes indicar isso; os signatários das duas
    # são somados (evitando duplicar o mesmo nome+data+hora).
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