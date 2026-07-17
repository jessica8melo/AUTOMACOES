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
# Verificação de assinatura eletrônica
# ---------------------------------------------------------------------------
# IMPORTANTE: isto é uma checagem TEXTUAL (procura o carimbo/texto que a
# plataforma de assinatura grava no PDF), não uma verificação criptográfica.
# Um PDF com assinatura digital real (ICP-Brasil, por exemplo) tem um objeto
# /Sig com /ByteRange no arquivo; esses documentos, em geral, só trazem um
# texto informativo ("Assinado eletronicamente através do sistema X").
# ---------------------------------------------------------------------------

# Frases que indicam que o documento foi assinado eletronicamente,
# uma por plataforma/formato conhecido. Adicione outras conforme aparecerem.
INDICADORES_ASSINATURA = [
    ("Aprovve", r"assinado\s+eletronicamente\s+atrav[eé]s\s+do\s+sistema\s+aprovve"),
    ("GOV.BR / ICP-Brasil", r"assinado\s+de\s+forma\s+eletr[oô]nica|verifique\s+em\s+https?://verificador\.iti\.gov\.br|assinatura\s+qualificada"),
    ("DocuSign", r"docusign"),
    ("Clicksign", r"clicksign"),
    ("Genérico", r"assinado\s+digitalmente|documento\s+assinado\s+eletronicamente"),
]

# Captura "NOME - CARGO - DD/MM/AAAA – HH:MM", formato usado pelo Aprovve
# (e possivelmente outras plataformas com o mesmo padrão de rodapé).
PADRAO_SIGNATARIO = re.compile(
    r"([A-ZÀ-Ú][A-ZÀ-Ú0-9\s/]+?)\s-\s(.+?)\s-\s(\d{2}/\d{2}/\d{4})\s[–-]\s(\d{2}:\d{2})"
)

# Número do processo/protocolo de assinatura (ex.: "sob o número 2026/007058")
PADRAO_NUMERO_PROCESSO = re.compile(
    r"sob\s+o\s+n[uú]mero\s+([\w./\-]+)", re.IGNORECASE
)


def verificar_assinatura(texto: str) -> dict:
    """
    Analisa o texto do PDF em busca de indícios de assinatura eletrônica.
    Devolve um dicionário:
        {
          "assinado": bool,
          "sistema": str ou None,         # plataforma identificada
          "numero_processo": str ou None, # nº de protocolo, se houver
          "signatarios": [                # lista de assinantes encontrados
              {"nome": ..., "cargo": ..., "data": ..., "hora": ...}, ...
          ],
        }
    Aviso: é uma checagem textual (procura o carimbo da plataforma), não uma
    verificação criptográfica da assinatura.
    """
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

    for nome, cargo, data, hora in PADRAO_SIGNATARIO.findall(texto):
        resultado["signatarios"].append({
            "nome": " ".join(nome.split()),
            "cargo": " ".join(cargo.split()),
            "data": data,
            "hora": hora,
        })

    # Se achou signatários com nome/data mas nenhuma frase-indicador bateu,
    # ainda assim considera assinado (documento pode usar uma plataforma
    # não mapeada em INDICADORES_ASSINATURA, mas com o mesmo padrão de rodapé)
    if resultado["signatarios"] and not resultado["assinado"]:
        resultado["assinado"] = True

    return resultado


# ---------------------------------------------------------------------------
# Função reutilizável (chamada pelo main.py e também usável via CLI)
# ---------------------------------------------------------------------------
def processar_pdf(caminho_pdf: str) -> dict:
    """
    Analisa o PDF em `caminho_pdf` e devolve um dicionário
    {campo: valor_encontrado_ou_None} para cada campo em CAMPOS_PROCURADOS,
    mais a chave especial "_assinatura" com o resultado de verificar_assinatura().
    Lança exceção se o arquivo não puder ser aberto/lido.
    """
    texto = extrair_texto(caminho_pdf)
    tabelas = extrair_tabelas(caminho_pdf)

    resultado = {}
    for campo in CAMPOS_PROCURADOS:
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


# ---------------------------------------------------------------------------
# Programa principal (uso via linha de comando, standalone)
# ---------------------------------------------------------------------------
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