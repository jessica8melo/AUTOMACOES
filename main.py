#!/usr/bin/env python3
"""
Ponto de entrada único.

Por enquanto o projeto foca só na EXTRAÇÃO DE DADOS dos documentos (a
parte dos fluxos/checklists de fluxos.py fica pausada para depois).

Etapas do script:

    1. Para cada arquivo, detecta o tipo real (PDF ou planilha) pelos
       bytes iniciais.
    2. Identifica qual DOCUMENTO ele é (Contrato, FQ415-075, Nota
       Técnica, Projeto Básico, Solicitação de Entrega — ver
       documentos.py), usando doc_types.py.
    3. Busca só os campos daquele documento (documentos.campos_do_documento)
       e só então chama pdfs.py/tabelas.py, que recebem essa lista de
       campos por parâmetro.

Uso:
    python main.py caminho/do/arquivo.pdf
    python main.py caminho/da/pasta
"""

import os
import sys

import doc_types
import documentos
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
# Processamento de um arquivo
# ---------------------------------------------------------------------------
def processar_arquivo(caminho: str) -> bool:
    """Detecta o tipo de `caminho`, identifica qual documento ele é (dentre
    documentos.DOCUMENTOS), busca só os campos desse documento e encaminha
    para o script correto. Devolve True se conseguiu processar, False caso
    contrário."""
    tipo = detectar_tipo(caminho)

    if tipo is None:
        print(f"[ERRO] Tipo de arquivo não reconhecido para '{caminho}'. "
              f"Suportados: PDF e XLSX/XLSM/XLS.")
        return False

    candidatos = documentos.listar_documentos()
    texto = extrair_texto_para_identificacao(caminho, tipo)
    tipo_documento = doc_types.identificar_documento(caminho, texto, candidatos)

    if tipo_documento is None:
        print(f"[ERRO] Não foi possível identificar qual documento o arquivo "
              f"'{caminho}' representa. Documentos reconhecidos: "
              f"{', '.join(candidatos)}.")
        return False

    campos_para_extrair = documentos.campos_do_documento(tipo_documento)

    print(f"[INFO] Documento identificado: '{tipo_documento}' "
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

    return True


def processar_pasta(caminho_pasta: str) -> None:
    """Processa, em ordem alfabética, todos os arquivos dentro de uma pasta
    (não entra em subpastas)."""
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
        if processar_arquivo(caminho_completo):
            sucessos += 1
        else:
            falhas += 1
        print()

    print("=" * 70)
    print(f"Resumo: {sucessos} arquivo(s) processado(s), {falhas} com erro, "
          f"de {len(arquivos)} arquivo(s) na pasta.")


def main():
    if len(sys.argv) != 2:
        print("Uso: python main.py caminho/do/arquivo.pdf")
        print("     python main.py caminho/da/pasta")
        sys.exit(1)

    caminho = sys.argv[1]

    if os.path.isdir(caminho):
        processar_pasta(caminho)
    elif os.path.isfile(caminho):
        if not processar_arquivo(caminho):
            sys.exit(1)
    else:
        print(f"[ERRO] Caminho não encontrado: '{caminho}'")
        sys.exit(1)


if __name__ == "__main__":
    main()
