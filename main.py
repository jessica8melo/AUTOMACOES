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
def extrair_texto(caminho_pdf: str) -> str:
    partes = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            partes.append(pagina.extract_text(layout=True) or "")
    return "\n".join(partes)


def extrair_tabelas(caminho_pdf: str):
    tabelas = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            tabelas.extend(pagina.extract_tables())
    return tabelas


def mapear_campos_em_tabela(tabela):
    """
    Varre uma tabela e devolve dois resultados:
      - pares_rotulo_valor: linhas soltas do tipo "Rótulo ... Valor"
      - cabecalho_valor: dicionário {texto_do_cabecalho: valor}, construído
        juntando cabeçalhos que se quebram em mais de uma linha e casando
        (na ordem em que aparecem) com a linha de dados correspondente.
    """
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

        # Linha "rótulo ... valor" isolada (ex.: "Valor total da solicitação | R$ 6.719,92")
        if tem_valor and len(celulas) == 2 and not parece_valor(celulas[0][1]):
            pares_rotulo_valor.append((celulas[0][1], celulas[1][1]))
            i += 1
            continue

        # Bloco de cabeçalho: linha(s) em que NENHUMA célula parece valor.
        # Uma linha de cabeçalho pode ter só 2 células (ex.: continuação de um
        # cabeçalho quebrado em duas linhas), então não filtramos por tamanho aqui.
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

            # Junta os textos de cabeçalho por posição de coluna original
            colunas = {}
            for celulas_linha in bloco_cabecalho:
                for col, texto in celulas_linha:
                    colunas[col] = (colunas.get(col, "") + " " + texto).strip()
            cabecalhos_em_ordem = [colunas[c] for c in sorted(colunas.keys())]

            # Linhas de dados seguintes (podem ser mais de um item)
            while j < total:
                dados = linhas_info[j]
                if not dados:
                    break
                dados_tem_valor = any(parece_valor(c) for _, c in dados)
                if not dados_tem_valor:
                    break  # começou outro bloco de cabeçalho
                if len(dados) == 2 and not parece_valor(dados[0][1]):
                    break  # é uma linha "rótulo: valor" isolada, não dado da tabela
                valores_em_ordem = [texto for _, texto in dados]
                for cabecalho, valor in zip(cabecalhos_em_ordem, valores_em_ordem):
                    # Não sobrescreve um valor já encontrado (mantém o primeiro item)
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


# ---------------------------------------------------------------------------
# Busca no texto corrido (fallback para campos que não estão em tabela)
# ---------------------------------------------------------------------------
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
        # Só aceita se a linha seguinte parecer um valor de fato (evita pegar
        # continuação de outro rótulo, como acontece em cabeçalhos de tabela)
        if valor and (parece_valor(valor) or len(valor.split()) <= 6):
            return valor

    return None


# ---------------------------------------------------------------------------
# Padrões especiais: alguns campos não aparecem como "Rótulo: valor", mas sim
# embutidos dentro de uma frase do documento. Cada campo pode ter uma lista
# de padrões alternativos, testados em ordem até um dar certo.
# ---------------------------------------------------------------------------
PADROES_ESPECIAIS = {
    "Data do fornecimento": [
        # Ex.: "firmada e assinada com essa empresa em 17 de outubro de 2025"
        # (\bem\s+ exige que "em" seja seguido de dígito, evitando casar com o "em" de "empresa")
        r"firmad[ao]\s+e\s+assinad[ao].{0,80}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4})",
        # Ex.: "contrato ... assinado em 17/10/2025" ou "assinada em 17 de outubro de 2025"
        r"assinad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
        # Ex.: "firmado em 17/10/2025"
        r"firmad[ao].{0,60}?\bem\s+(?=\d)(\d{1,2}\s+de\s+[^\W\d_]+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{2,4})",
    ],
    "DGCO nº": [
        # Ex.: "contrato DGCO nº 02732-2025, referente..."
        r"DGCO\s*n[ºo°]?\s*([\d./\-]+)",
    ],
    "OC Master nº": [
        # Ex.: "OC Master nº 196371"
        r"OC\s*Master\s*n[ºo°]?\s*([\d./\-]+)",
    ],
}


def buscar_com_padroes_especiais(texto: str, campo: str):
    padroes = PADROES_ESPECIAIS.get(campo)
    if not padroes:
        return None
    for padrao in padroes:
        match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
        if match:
            return " ".join(match.group(1).split())
    return None


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python buscar_campos_pdf_v2.py caminho/do/arquivo.pdf")
        sys.exit(1)

    caminho_pdf = sys.argv[1]

    try:
        texto = extrair_texto(caminho_pdf)
        tabelas = extrair_tabelas(caminho_pdf)
    except Exception as erro:
        print(f"[ERRO] Não foi possível abrir/ler o arquivo '{caminho_pdf}': {erro}")
        sys.exit(1)

    print(f"Documento analisado: {caminho_pdf}\n")

    for campo in CAMPOS_PROCURADOS:
        valor = buscar_em_tabelas(tabelas, campo)
        if not valor:
            valor = buscar_em_texto(texto, campo)
        if not valor:
            valor = buscar_com_padroes_especiais(texto, campo)

        if valor:
            valor_limpo = " ".join(valor.split())
            print(f"[SUCESSO] Campo '{campo}' encontrado. Valor: {valor_limpo}")
        else:
            print(f"[ERRO] Não foi possível encontrar o campo '{campo}' no documento.")


if __name__ == "__main__":
    main()