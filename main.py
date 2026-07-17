#!/usr/bin/env python3
"""
Ponto de entrada único: recebe o caminho de um arquivo, identifica o tipo
(PDF ou planilha) e encaminha para o script correto:
    - PDF           -> pdfs.py     (processar_pdf)
    - XLSX/XLS/XLSM -> tabelas.py  (processar_xlsx)

A identificação do tipo usa a extensão do arquivo como primeira pista, mas
confirma pelos bytes iniciais do arquivo (assinatura/"magic number"), para
não se enganar se o arquivo estiver com a extensão errada.

Uso:
    python main.py caminho/do/arquivo.pdf
    python main.py caminho/do/arquivo.xlsx
    python main.py caminho/da/pasta          (processa todos os arquivos da pasta)
"""

import sys
import os

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


def processar_arquivo(caminho: str) -> bool:
    """Detecta o tipo de `caminho` e encaminha para o script correto.
    Devolve True se conseguiu processar, False caso contrário."""
    tipo = detectar_tipo(caminho)

    if tipo == "pdf":
        print("[INFO] Tipo detectado: PDF -> encaminhando para pdfs.py\n")
        try:
            resultado = pdfs.processar_pdf(caminho)
        except Exception as erro:
            print(f"[ERRO] Não foi possível abrir/ler o arquivo '{caminho}': {erro}")
            return False
        pdfs.imprimir_resultado(caminho, resultado)
        return True

    elif tipo == "xlsx":
        print("[INFO] Tipo detectado: planilha -> encaminhando para tabelas.py\n")
        try:
            resultado = tabelas.processar_xlsx(caminho)
        except Exception as erro:
            print(f"[ERRO] Não foi possível abrir o arquivo '{caminho}': {erro}")
            return False
        tabelas.imprimir_resultado(caminho, resultado)
        return True

    else:
        print(f"[ERRO] Tipo de arquivo não reconhecido para '{caminho}'. "
              f"Suportados: PDF e XLSX/XLSM/XLS.")
        return False


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
    if len(sys.argv) < 2:
        print("Uso: python main.py caminho/do/arquivo (.pdf ou .xlsx) OU caminho/da/pasta")
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