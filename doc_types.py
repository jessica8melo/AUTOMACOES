#!/usr/bin/env python3
"""
Identifica QUAL documento um arquivo representa (Contrato, Nota Técnica,
FQ415-075, Projeto Básico, Solicitação de Entrega etc.), usando
marcadores no nome do arquivo e no texto do seu conteúdo.

Este módulo não define a lista de documentos reconhecidos nem seus
campos — isso está em documentos.py. Aqui só respondemos: "dentre os
documentos passados em `candidatos` (normalmente
documentos.listar_documentos()), qual deles é este arquivo?".

Cada categoria de documento tem:
  - "aliases": as chaves usadas em fluxos.FLUXOS que se referem a ela
    (podem variar ligeiramente de fluxo para fluxo, ex.: "Contrato" vs.
    "Contrato / Aditivo" — ambas apontam para a mesma categoria aqui).
  - "nome_arquivo": padrões (regex) que, se baterem no NOME do arquivo,
    são um indício forte do tipo de documento.
  - "texto": padrões (regex) que, se baterem no TEXTO/conteúdo do
    arquivo, são um indício mais fraco (mas ainda útil, especialmente
    quando o nome do arquivo é genérico, ex.: um uuid.pdf).

Os padrões abaixo são aplicados sobre texto já normalizado (minúsculo,
sem acento, sem pontuação — ver `normalizar()` em pdfs.py), por isso são
escritos sem acentuação.
"""

import os
import re

from pdfs import normalizar

PESO_NOME_ARQUIVO = 3
PESO_TEXTO = 1

# Quantos padrões "fracos" (texto OU nome_arquivo_fraco) distintos precisam
# bater para um candidato ser aceito quando não há nenhum acerto de nome de
# arquivo FORTE. Uma única menção solta do código/rótulo (ex.: uma
# referência de passagem a "FQ415-075" dentro de uma Ordem de Serviço que
# na verdade é outro tipo de documento) não deve, sozinha, classificar o
# arquivo inteiro.
MIN_ACERTOS_TEXTO = 2

MARCADORES = {
    "contrato": {
        "aliases": ["Contrato", "Contrato / Aditivo"],
        "nome_arquivo": [r"contrat", r"aditivo", r"registro de preco"],
        # Siglas curtas (2-4 letras) são um indício FRACO pelo nome do
        # arquivo: "arp", "pb", "nt", "dra" etc. podem aparecer soltas em
        # QUALQUER nome de arquivo só como referência a outro documento
        # (ex.: uma planilha de saldo cujo nome cita "NT 2023-0574" sem
        # ela mesma ser a Nota Técnica). Por isso entram com o mesmo peso
        # baixo de um padrão de texto (PESO_TEXTO), não de PESO_NOME_ARQUIVO
        # — assim elas só decidem a classificação quando reforçadas por
        # outro indício (outro padrão de nome_arquivo_fraco OU de texto).
        "nome_arquivo_fraco": [r"\barp\b"],
        "texto": [
            r"instrumento (particular )?de contrato",
            r"\baditivo\b",
            r"ata de registro de preco",
            r"contratante.{0,40}contratad[ao]",
            r"clausula (primeira|segunda|terceira)",
        ],
    },
    "projeto_basico": {
        "aliases": ["Projeto Básico"],
        "nome_arquivo": [r"projeto\s*basic"],
        "nome_arquivo_fraco": [r"\bpb\b"],
        "texto": [r"projeto basico"],
    },
    "nota_tecnica": {
        "aliases": ["Nota Técnica"],
        "nome_arquivo": [r"nota\s*tecnic"],
        "nome_arquivo_fraco": [r"\bnt\b"],
        "texto": [r"nota tecnica", r"numero da nota tecnica"],
    },
    "fq415_075": {
        "aliases": ["FQ415-075"],
        "nome_arquivo": [r"fq[\s_-]*415[\s_-]*0?75"],
        "texto": [r"fq\s*415[\s_-]*0?75"],
    },
    "acc_master": {
        "aliases": ["ACC Master"],
        "nome_arquivo": [r"acc[\s_-]*master"],
        "texto": [r"acc\s*master", r"numero da oc master"],
    },
    "solicitacao_entrega": {
        "aliases": ["Solicitação de Entrega"],
        "nome_arquivo": [r"solicitacao\s*de\s*entrega"],
        "texto": [r"solicitacao de entrega de bens", r"solicitacao de entrega"],
    },
    "ata_condominio": {
        "aliases": ["Ata do Condomínio"],
        "nome_arquivo": [r"ata.*condomini"],
        "texto": [r"ata d[ao] condomini", r"assembleia.*condomini"],
    },
    "boleto_condominio": {
        "aliases": ["Boleto do Condomínio"],
        "nome_arquivo": [r"boleto.*condomini"],
        "texto": [r"boleto d[ao] condomini", r"condomini.*boleto"],
    },
    "doc_referencia_area": {
        "aliases": ["Documento de Referência da Área"],
        "nome_arquivo": [r"referencia.*area"],
        "nome_arquivo_fraco": [r"\bdra\b"],
        "texto": [r"documento de referencia da area"],
    },
    "fq412_034": {
        "aliases": ["FQ412-034"],
        "nome_arquivo": [r"fq[\s_-]*412[\s_-]*0?34"],
        "texto": [r"fq\s*412[\s_-]*0?34"],
    },
    "fq412_035": {
        "aliases": ["FQ412-035"],
        "nome_arquivo": [r"fq[\s_-]*412[\s_-]*0?35"],
        "texto": [r"fq\s*412[\s_-]*0?35"],
    },
}


def _marcadores_para_chave(chave: str) -> dict:
    """Encontra a categoria de MARCADORES cujo apelido bate com `chave`
    (a chave de documento tal como usada em fluxos.FLUXOS).

    Usa comparação EXATA (case-insensitive) em vez da comparação
    aproximada `parecido()`: as chaves em fluxos.FLUXOS são strings
    fixas e conhecidas (não texto livre extraído de documento), e
    códigos curtos como "FQ412-034"/"FQ412-035"/"FQ415-075" são parecidos
    demais entre si para a comparação aproximada — ela classificava um
    errado como o outro.
    """
    chave_normalizada = normalizar(chave)
    for categoria in MARCADORES.values():
        for apelido in categoria["aliases"]:
            if normalizar(apelido) == chave_normalizada:
                return categoria
    return {}


def identificar_documento(caminho: str, texto: str, candidatos: list) -> str:
    """
    Devolve qual, dentre as chaves em `candidatos` (os documentos que o
    fluxo atual espera — ver fluxos.documentos_do_fluxo), melhor
    corresponde ao arquivo em `caminho`, combinando indícios do nome do
    arquivo com indícios do texto/conteúdo.

    Cada categoria tem 3 tipos de indício:
      - "nome_arquivo": marcador FORTE e específico no nome do arquivo
        (ex.: "fq412034", "nota tecnica", "projeto basico"). Peso
        PESO_NOME_ARQUIVO.
      - "nome_arquivo_fraco": sigla curta e ambígua no nome do arquivo
        (ex.: "nt", "pb", "arp", "dra"), que pode aparecer só como
        referência solta dentro do nome de um documento de OUTRO tipo
        (ex.: uma planilha de saldo chamada "...SALDO NT 2023-0574..."
        não é, ela mesma, a Nota Técnica). Peso PESO_TEXTO (o mesmo de um
        padrão de texto), para nunca decidir a classificação sozinha.
      - "texto": padrão no conteúdo do arquivo. Peso PESO_TEXTO.

    LIMIAR MÍNIMO DE CONFIANÇA: um candidato só é aceito se tiver pelo
    menos 1 acerto de "nome_arquivo" (forte), OU pelo menos
    `MIN_ACERTOS_TEXTO` acertos somando "nome_arquivo_fraco" + "texto"
    (indícios fracos, mas reforçados um pelo outro). Um único indício
    fraco isolado (ex.: uma sigla solta no nome, ou uma menção solta no
    texto) nunca é suficiente para classificar o arquivo inteiro.

    Devolve None se nenhum candidato tiver indício suficiente (nesse
    caso, quem chama deve avisar o usuário em vez de adivinhar).
    """
    nome_normalizado = normalizar(os.path.basename(caminho))
    texto_normalizado = normalizar(texto) if texto else ""

    melhor_chave = None
    melhor_pontuacao = 0

    for chave in candidatos:
        categoria = _marcadores_para_chave(chave)
        if not categoria:
            continue

        acertos_nome_forte = sum(
            1 for padrao in categoria.get("nome_arquivo", [])
            if re.search(padrao, nome_normalizado)
        )
        acertos_nome_fraco = sum(
            1 for padrao in categoria.get("nome_arquivo_fraco", [])
            if re.search(padrao, nome_normalizado)
        )
        acertos_texto = sum(
            1 for padrao in categoria.get("texto", [])
            if re.search(padrao, texto_normalizado)
        )
        acertos_fracos = acertos_nome_fraco + acertos_texto

        if acertos_nome_forte == 0 and acertos_fracos < MIN_ACERTOS_TEXTO:
            continue  # indício insuficiente: não confia neste candidato

        pontuacao = acertos_nome_forte * PESO_NOME_ARQUIVO + acertos_fracos * PESO_TEXTO

        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_chave = chave

    return melhor_chave if melhor_pontuacao > 0 else None