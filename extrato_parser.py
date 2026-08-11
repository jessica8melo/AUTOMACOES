"""
Extrator de extrato de conta corrente do Banco do Brasil (PDF -> dados estruturados)
=====================================================================================

Lê um PDF (ou uma pasta com vários PDFs) de "Consultas - Extrato de conta corrente"
do BB e extrai:
  - Agência, número da conta e titular
  - Dia (ou período) do extrato
  - Lançamentos: histórico, documento, valor e se é Recebimento (C) ou Pagamento (D)

A extração usa as TABELAS do PDF (pdfplumber), não o texto corrido — isso deixa o
parsing muito mais confiável, inclusive em extratos com muitas linhas e múltiplas
páginas.

Uso:
    # um único PDF
    python extrato_bb_parser.py extrato.pdf

    # uma pasta inteira (varre todos os .pdf, inclusive em subpastas)
    python extrato_bb_parser.py --dir caminho/da/pasta

    # salvando os resultados
    python extrato_bb_parser.py --dir caminho/da/pasta --csv saida.csv --json saida.json
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path

import pdfplumber

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
LOTE_PREFIX_RE = re.compile(r"^\d{2,4}\s+")

# Larguras de colunas da tabela "Lançamentos" no layout do BB:
# Dt.balancete | Dt.movimento | Ag.origem | Lote | Histórico | Documento | Valor R$ | Saldo
N_COLS_LANCAMENTOS = 8


def _limpar(valor):
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", valor.replace("\n", " ")).strip()


def _extrair_linhas_tabela(pdf_path: str):
    """Percorre todas as páginas do PDF e junta as linhas da tabela de lançamentos
    (ela pode continuar em várias páginas, sem repetir o cabeçalho)."""
    linhas = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for tabela in page.extract_tables():
                if not tabela or len(tabela[0]) != N_COLS_LANCAMENTOS:
                    continue
                for row in tabela:
                    primeira = _limpar(row[0])
                    if primeira.startswith("Dt."):
                        continue  # linha de cabeçalho
                    linhas.append(row)
    return linhas


def _cabecalho(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    linhas_txt = [l.strip() for l in texto.split("\n")]

    agencia = conta_raw = periodo = None
    for l in linhas_txt:
        if l.startswith("Agência"):
            agencia = l.replace("Agência", "").strip()
        elif l.startswith("Conta corrente"):
            conta_raw = l.replace("Conta corrente", "").strip()
        elif l.startswith("Período do extrato"):
            periodo = l.replace("Período do extrato", "").strip()

    m = re.match(r"^([\d.\-]+)\s*(.*)$", conta_raw or "")
    numero_conta = m.group(1) if m else conta_raw
    titular = m.group(2).strip() if m and m.group(2) else None

    datas = re.findall(r"\d{2}\s*/\s*\d{2}\s*/\s*\d{4}", periodo or "")
    datas = [d.replace(" ", "") for d in datas]
    dia_extrato = datas[0] if len(datas) == 2 and datas[0] == datas[1] else None

    return {
        "agencia": agencia,
        "numero_conta": numero_conta,
        "titular": titular,
        "dia_extrato": dia_extrato,
        "periodo": periodo,
    }


def parse_extrato(pdf_path: str) -> dict:
    """Lê um PDF de extrato BB e devolve um dicionário estruturado."""
    dados = _cabecalho(pdf_path)
    linhas = _extrair_linhas_tabela(pdf_path)

    lancamentos = []
    for row in linhas:
        dt_balancete = _limpar(row[0])

        if dt_balancete and DATE_RE.match(dt_balancete):
            historico = LOTE_PREFIX_RE.sub("", _limpar(row[4]))
            documento = _limpar(row[5]) or None

            valor_col = _limpar(row[6])   # "Valor R$" -> lançamentos normais
            saldo_col = _limpar(row[7])   # "Saldo"    -> saldo anterior / saldo do dia
            bruto = valor_col or saldo_col

            tokens = bruto.split()
            tipo = tokens[-1] if tokens and tokens[-1] in ("C", "D") else None
            valor = tokens[-2] if tipo and len(tokens) >= 2 else (tokens[0] if tokens else None)

            lancamentos.append({
                "data": dt_balancete,
                "historico": historico,
                "documento": documento,
                "valor": valor,
                "natureza": ("Recebimento" if tipo == "C" else "Pagamento") if tipo else None,
                "tipo_registro": "lancamento" if documento else "saldo",
            })
        else:
            # linha de continuação: acrescenta ao histórico do lançamento anterior
            complemento = _limpar(row[4])
            if complemento and lancamentos:
                lancamentos[-1]["historico"] = (lancamentos[-1]["historico"] + " " + complemento).strip()

    dados["lancamentos"] = lancamentos
    dados["arquivo"] = Path(pdf_path).name
    return dados


def parse_pasta(dir_path: str, recursivo: bool = True) -> list:
    """Lê todos os PDFs de uma pasta e devolve uma lista de dicionários (um por arquivo)."""
    pasta = Path(dir_path)
    padrao = "**/*.pdf" if recursivo else "*.pdf"
    arquivos = sorted(pasta.glob(padrao))

    if not arquivos:
        print(f"Nenhum PDF encontrado em: {dir_path}", file=sys.stderr)

    resultados = []
    for arq in arquivos:
        try:
            resultados.append(parse_extrato(str(arq)))
        except Exception as e:
            print(f"Falha ao processar '{arq.name}': {e}", file=sys.stderr)
    return resultados


def salvar_csv(lista_extratos: list, caminho: str) -> None:
    campos = ["arquivo", "agencia", "numero_conta", "dia_extrato",
              "data", "historico", "documento", "valor", "natureza", "tipo_registro"]
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for extrato in lista_extratos:
            for lc in extrato["lancamentos"]:
                writer.writerow({
                    "arquivo": extrato["arquivo"],
                    "agencia": extrato["agencia"],
                    "numero_conta": extrato["numero_conta"],
                    "dia_extrato": extrato["dia_extrato"],
                    **lc,
                })


def imprimir_resumo(dados: dict) -> None:
    print(f"\nArquivo: {dados['arquivo']}")
    print(f"Agência: {dados['agencia']}  |  Conta: {dados['numero_conta']}"
          + (f" ({dados['titular']})" if dados.get("titular") else ""))
    print(f"Dia do extrato: {dados['dia_extrato'] or dados['periodo']}")
    print("-" * 90)
    for lc in dados["lancamentos"]:
        marca = "[SALDO]" if lc["tipo_registro"] == "saldo" else f"[{lc['natureza']}]"
        doc = lc["documento"] or "-"
        print(f"{lc['data']}  {marca:<14} R$ {lc['valor']:<12} doc: {doc:<22} {lc['historico']}")
    print("-" * 90)


def main():
    parser = argparse.ArgumentParser(description="Extrai lançamentos de extrato(s) BB em PDF")
    grupo = parser.add_mutually_exclusive_group(required=True)
    grupo.add_argument("pdf", nargs="?", help="Caminho de um único arquivo PDF")
    grupo.add_argument("--dir", help="Caminho de uma pasta com vários PDFs")
    parser.add_argument("--sem-recursivo", action="store_true",
                         help="Com --dir, não entra em subpastas (por padrão entra)")
    parser.add_argument("--csv", help="Caminho para salvar os lançamentos em CSV", default=None)
    parser.add_argument("--json", help="Caminho para salvar os dados completos em JSON", default=None)
    parser.add_argument("--silencioso", action="store_true", help="Não imprime o resumo no terminal")
    args = parser.parse_args()

    if args.dir:
        if not Path(args.dir).is_dir():
            print(f"Pasta não encontrada: {args.dir}", file=sys.stderr)
            sys.exit(1)
        resultados = parse_pasta(args.dir, recursivo=not args.sem_recursivo)
    else:
        if not Path(args.pdf).exists():
            print(f"Arquivo não encontrado: {args.pdf}", file=sys.stderr)
            sys.exit(1)
        resultados = [parse_extrato(args.pdf)]

    if not args.silencioso:
        for dados in resultados:
            imprimir_resumo(dados)
        total = sum(len(d["lancamentos"]) for d in resultados)
        print(f"\n{len(resultados)} arquivo(s) processado(s), {total} linha(s) extraída(s) no total.")

    if args.csv:
        salvar_csv(resultados, args.csv)
        print(f"\nCSV salvo em: {args.csv}")

    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=2)
        print(f"JSON salvo em: {args.json}")


if __name__ == "__main__":
    main()