#!/usr/bin/env python3
"""
Ponto de entrada único.

Etapas do script:

    1. Pergunta/recebe qual dos fluxos (checklists) de fluxos.py está
       sendo aplicado (ex.: FQ415-031 - OC Padrão com Contrato).
    2. Para cada arquivo, detecta o tipo real (PDF ou planilha) pelos
       bytes iniciais, e também qual DOCUMENTO especifico do fluxo aquele
       arquivo representa (Contrato, Nota Técnica, ACC Master etc.),
       usando doc_types.py.
    3. Busca só os campos daquela combinação fluxo+documento (em vez de
       uma lista fixa única) e só então chama pdfs.py/tabelas.py, que
       agora recebem essa lista de campos por parâmetro.

Uso:
    python main.py --fluxo FQ415-031 caminho/do/arquivo.pdf
    python main.py --fluxo FQ415-031 caminho/da/pasta
    python main.py caminho/do/arquivo.pdf   (pergunta o fluxo interativamente)
"""

import argparse
import os
import sys

import doc_types
import fluxos
import pdfs
import tabelas

EXTENSOES_PDF = {".pdf"}
EXTENSOES_XLSX = {".xlsx", ".xlsm", ".xls"}


def assinatura_do_arquivo(caminho: str) -> bytes:
    """Lê os primeiros bytes do arquivo para conferir sua assinatura real."""
    with open(caminho, "rb") as f:
        return f.read(8)


def detectar_tipo(caminho: str) -> str:
    """
    Devolve 'pdf', 'xlsx' ou None (tipo não reconhecido).
    Usa a extensão como primeira pista e confirma pela assinatura de bytes:
      - PDF começa com  %PDF
      - XLSX/XLSM são arquivos ZIP e começam com  PK\\x03\\x04
      - XLS (formato antigo, binário OLE) começa com  \\xD0\\xCF\\x11\\xE0
    """
    _, extensao = os.path.splitext(caminho.lower())

    try:
        assinatura = assinatura_do_arquivo(caminho)
    except Exception:
        assinatura = b""

    eh_pdf_pela_assinatura = assinatura.startswith(b"%PDF")
    eh_zip_pela_assinatura = assinatura.startswith(b"PK\x03\x04")
    eh_ole_pela_assinatura = assinatura.startswith(b"\xd0\xcf\x11\xe0")

    if extensao in EXTENSOES_PDF or eh_pdf_pela_assinatura:
        return "pdf"

    if extensao in EXTENSOES_XLSX or eh_zip_pela_assinatura or eh_ole_pela_assinatura:
        return "xlsx"

    return None


# ---------------------------------------------------------------------------
# Escolha do fluxo
# ---------------------------------------------------------------------------
def escolher_fluxo_interativamente() -> str:
    """Lista os fluxos disponíveis e pede ao usuário para escolher um pelo
    número ou pelo id (ex.: FQ415-031)."""
    lista = fluxos.listar_fluxos()

    print("Qual fluxo (checklist) está sendo aplicado?\n")
    for i, (id_fluxo, nome_fluxo) in enumerate(lista, start=1):
        print(f"  {i}. {id_fluxo} — {nome_fluxo}")
    print()

    while True:
        escolha = input("Digite o número ou o id do fluxo: ").strip()

        if escolha.isdigit() and 1 <= int(escolha) <= len(lista):
            return lista[int(escolha) - 1][0]

        for id_fluxo, _ in lista:
            if escolha.upper() == id_fluxo.upper():
                return id_fluxo

        print(f"[ERRO] '{escolha}' não é um número ou id de fluxo válido. Tente novamente.\n")


def validar_fluxo(id_fluxo: str) -> str:
    """Confere se `id_fluxo` existe em fluxos.FLUXOS (sem diferenciar
    maiúsculas/minúsculas). Devolve o id no formato correto ou None."""
    for id_valido, _ in fluxos.listar_fluxos():
        if id_fluxo.upper() == id_valido.upper():
            return id_valido
    return None


# ---------------------------------------------------------------------------
# Extração de texto (para identificar o documento) por tipo de arquivo
# ---------------------------------------------------------------------------
def extrair_texto_para_identificacao(caminho: str, tipo: str) -> str:
    """Devolve um texto representativo do conteúdo do arquivo, usado só
    para RECONHECER qual documento ele é (não para extrair campos)."""
    try:
        if tipo == "pdf":
            return pdfs.extrair_texto(caminho)
        elif tipo == "xlsx":
            return tabelas.extrair_texto_xlsx(caminho)
    except Exception as erro:
        print(f"[AVISO] Não foi possível ler o conteúdo de '{caminho}' para identificação: {erro}",
              file=sys.stderr)
    return ""


# ---------------------------------------------------------------------------
# Processamento de um arquivo dentro de um fluxo já escolhido
# ---------------------------------------------------------------------------
def processar_arquivo(caminho: str, id_fluxo: str) -> bool:
    """Detecta o tipo de `caminho`, identifica qual documento do fluxo
    `id_fluxo` ele é, busca só os campos dessa combinação fluxo+documento
    e encaminha para o script correto. Devolve True se conseguiu
    processar, False caso contrário."""
    tipo = detectar_tipo(caminho)

    if tipo is None:
        print(f"[ERRO] Tipo de arquivo não reconhecido para '{caminho}'. "
              f"Suportados: PDF e XLSX/XLSM/XLS.")
        return False

    documentos_esperados = fluxos.documentos_do_fluxo(id_fluxo)
    texto = extrair_texto_para_identificacao(caminho, tipo)
    tipo_documento = doc_types.identificar_documento(caminho, texto, list(documentos_esperados.keys()))

    if tipo_documento is None:
        print(f"[ERRO] Não foi possível identificar qual documento do fluxo '{id_fluxo}' "
              f"o arquivo '{caminho}' representa. Documentos esperados nesse fluxo: "
              f"{', '.join(documentos_esperados.keys())}.")
        return False

    itens_checklist = fluxos.campos_do_documento(id_fluxo, tipo_documento)
    campos_para_extrair = [item for item in itens_checklist if not fluxos.eh_item_assinatura(item)]
    itens_assinatura = [item for item in itens_checklist if fluxos.eh_item_assinatura(item)]

    print(f"[INFO] Fluxo: {id_fluxo} — Documento identificado: '{tipo_documento}' "
          f"({'PDF' if tipo == 'pdf' else 'planilha'}) -> "
          f"encaminhando para {'pdfs.py' if tipo == 'pdf' else 'tabelas.py'}\n")

    if tipo == "pdf":
        try:
            resultado = pdfs.processar_pdf(caminho, campos_para_extrair)
        except Exception as erro:
            print(f"[ERRO] Não foi possível abrir/ler o arquivo '{caminho}': {erro}")
            return False
        pdfs.imprimir_resultado(caminho, resultado)

    else:  # xlsx
        try:
            resultado = tabelas.processar_xlsx(caminho, campos_para_extrair)
        except Exception as erro:
            print(f"[ERRO] Não foi possível abrir o arquivo '{caminho}': {erro}")
            return False
        tabelas.imprimir_resultado(caminho, resultado)

    for item in itens_assinatura:
        print(f"[CHECKLIST] {fluxos.descricao_item_assinatura(item)} "
              f"(ver seção [ASSINATURA] acima, se o arquivo for um PDF)")

    return True


def processar_pasta(caminho_pasta: str, id_fluxo: str) -> None:
    """Processa, em ordem alfabética, todos os arquivos dentro de uma pasta
    (não entra em subpastas), todos sob o mesmo fluxo `id_fluxo`."""
    arquivos = sorted(
        f for f in os.listdir(caminho_pasta)
        if os.path.isfile(os.path.join(caminho_pasta, f)) and not f.startswith(".")
    )

    if not arquivos:
        print(f"[ERRO] Nenhum arquivo encontrado em '{caminho_pasta}'.")
        return

    sucessos = 0
    falhas = 0

    for nome_arquivo in arquivos:
        caminho_completo = os.path.join(caminho_pasta, nome_arquivo)
        print("=" * 70)
        print(f"Arquivo: {nome_arquivo}")
        print("=" * 70)
        if processar_arquivo(caminho_completo, id_fluxo):
            sucessos += 1
        else:
            falhas += 1
        print()

    print("=" * 70)
    print(f"Resumo: {sucessos} arquivo(s) processado(s), {falhas} com erro, "
          f"de {len(arquivos)} arquivo(s) na pasta.")


def main():
    parser = argparse.ArgumentParser(
        description="Identifica o fluxo/checklist e os documentos de cada arquivo, "
                    "e extrai só os campos esperados para cada combinação fluxo+documento."
    )
    parser.add_argument("caminho", help="Caminho de um arquivo (PDF/XLSX) ou de uma pasta")
    parser.add_argument(
        "--fluxo",
        help="Id do fluxo/checklist a aplicar (ex.: FQ415-031). "
             "Se omitido, o script pergunta interativamente.",
    )
    args = parser.parse_args()

    if args.fluxo:
        id_fluxo = validar_fluxo(args.fluxo)
        if id_fluxo is None:
            print(f"[ERRO] Fluxo '{args.fluxo}' não existe. Fluxos disponíveis: "
                  f"{', '.join(id for id, _ in fluxos.listar_fluxos())}.")
            sys.exit(1)
    else:
        id_fluxo = escolher_fluxo_interativamente()

    print(f"\n[INFO] Fluxo em uso: {id_fluxo} — {dict(fluxos.listar_fluxos())[id_fluxo]}\n")

    caminho = args.caminho

    if os.path.isdir(caminho):
        processar_pasta(caminho, id_fluxo)
    elif os.path.isfile(caminho):
        if not processar_arquivo(caminho, id_fluxo):
            sys.exit(1)
    else:
        print(f"[ERRO] Caminho não encontrado: '{caminho}'")
        sys.exit(1)


if __name__ == "__main__":
    main()