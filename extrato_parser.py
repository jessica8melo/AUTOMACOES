"""
Extrator de extrato de conta corrente do Banco do Brasil (PDF -> dados estruturados)
=====================================================================================

Lê o PDF de "Consultas - Extrato de conta corrente" do BB e extrai:
  - Agência
  - Número da conta / titular
  - Dia (ou período) do extrato
  - Lançamentos: histórico, documento, valor e se é Recebimento (C) ou Pagamento (D)

Uso:
    python extrato_parser.py caminho/do/extrato.pdf
    python extrato_parser.py caminho/do/extrato.pdf --csv saida.csv
    python extrato_parser.py caminho/do/extrato.pdf --json saida.json
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path

import pdfplumber

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
DOC_RE = re.compile(r"^\d{2,3}(\.\d{2,3}){2,}$")


def _extrair_texto(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        paginas = [p.extract_text() or "" for p in pdf.pages]
    return "\n".join(paginas)


def parse_extrato(pdf_path: str) -> dict:
    """Lê o PDF do extrato BB e devolve um dicionário estruturado."""
    texto = _extrair_texto(pdf_path)
    linhas = [l.strip() for l in texto.split("\n")]

    agencia = conta_raw = periodo = None
    for l in linhas:
        if l.startswith("Agência"):
            agencia = l.replace("Agência", "").strip()
        elif l.startswith("Conta corrente"):
            conta_raw = l.replace("Conta corrente", "").strip()
        elif l.startswith("Período do extrato"):
            periodo = l.replace("Período do extrato", "").strip()

    # Separa "200000-8BB T SERVICOS S.A." em número da conta + titular
    m = re.match(r"^([\d.\-]+)\s*(.*)$", conta_raw or "")
    numero_conta = m.group(1) if m else conta_raw
    titular = m.group(2).strip() if m and m.group(2) else None

    # Datas do período: "de DD / MM / AAAA até DD / MM / AAAA"
    datas = re.findall(r"\d{2}\s*/\s*\d{2}\s*/\s*\d{4}", periodo or "")
    datas = [d.replace(" ", "") for d in datas]
    dia_extrato = datas[0] if len(datas) == 2 and datas[0] == datas[1] else None

    try:
        inicio = linhas.index(next(l for l in linhas if l.startswith("Lançamentos")))
        fim = linhas.index(next(l for l in linhas if l.startswith("Valores bloqueados")))
    except StopIteration:
        raise ValueError("Não foi possível localizar a seção de lançamentos no PDF.")

    bloco = linhas[inicio + 2:fim]  # pula título "Lançamentos" e cabeçalho da tabela

    lancamentos = []
    for raw in bloco:
        if not raw:
            continue
        tokens = raw.split()

        if DATE_RE.match(tokens[0]):
            data = tokens[0]
            resto = tokens[1:]

            tipo = resto[-1] if resto and resto[-1] in ("C", "D") else None
            valor = resto[-2] if tipo else None
            corte = -2 if tipo else len(resto)

            documento = None
            if tipo and len(resto) >= 3 and DOC_RE.match(resto[-3]):
                documento = resto[-3]
                corte = -3

            meio = resto[:corte]
            # descarta os códigos numéricos iniciais (agência de origem / lote)
            i = 0
            while i < len(meio) and meio[i].isdigit():
                i += 1
            historico = " ".join(meio[i:]).strip()

            lancamentos.append({
                "data": data,
                "historico": historico,
                "documento": documento,
                "valor": valor,
                "natureza": ("Recebimento" if tipo == "C" else "Pagamento") if tipo else None,
                "tipo_registro": "lancamento" if documento else "saldo",
            })
        else:
            # linha de continuação (ex.: detalhe do TED, "Cobrança referente a...")
            if lancamentos:
                lancamentos[-1]["historico"] = (lancamentos[-1]["historico"] + " " + raw).strip()

    return {
        "agencia": agencia,
        "numero_conta": numero_conta,
        "titular": titular,
        "dia_extrato": dia_extrato,
        "periodo": periodo,
        "lancamentos": lancamentos,
    }


def salvar_csv(dados: dict, caminho: str) -> None:
    campos = ["data", "historico", "documento", "valor", "natureza", "tipo_registro"]
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        writer.writerows(dados["lancamentos"])


def imprimir_resumo(dados: dict) -> None:
    print(f"Agência: {dados['agencia']}")
    print(f"Conta: {dados['numero_conta']}" + (f" ({dados['titular']})" if dados['titular'] else ""))
    print(f"Dia do extrato: {dados['dia_extrato'] or dados['periodo']}")
    print("\nLançamentos:")
    print("-" * 90)
    for lc in dados["lancamentos"]:
        marca = "[SALDO]" if lc["tipo_registro"] == "saldo" else f"[{lc['natureza']}]"
        doc = lc["documento"] or "-"
        print(f"{lc['data']}  {marca:<14} R$ {lc['valor']:<12} doc: {doc:<22} {lc['historico']}")
    print("-" * 90)


def main():
    parser = argparse.ArgumentParser(description="Extrai lançamentos de extrato BB em PDF")
    parser.add_argument("pdf", help="Caminho do arquivo PDF do extrato")
    parser.add_argument("--csv", help="Caminho para salvar os lançamentos em CSV", default=None)
    parser.add_argument("--json", help="Caminho para salvar os dados completos em JSON", default=None)
    args = parser.parse_args()

    if not Path(args.pdf).exists():
        print(f"Arquivo não encontrado: {args.pdf}", file=sys.stderr)
        sys.exit(1)

    dados = parse_extrato(args.pdf)
    imprimir_resumo(dados)

    if args.csv:
        salvar_csv(dados, args.csv)
        print(f"\nCSV salvo em: {args.csv}")

    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        print(f"JSON salvo em: {args.json}")


if __name__ == "__main__":
    main()
