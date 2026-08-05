#!/usr/bin/env python3
"""
Le o conteudo de arquivos "IEDDEB*.ret" (retornos CNAB240 de conta
corrente do BB), identifica o tipo de conta a partir de posicoes fixas
no registro-header (primeira linha do arquivo, registro tipo 0) e a
data de lancamento a partir do primeiro registro de detalhe (segmento
E), e renomeia o arquivo seguindo as regras abaixo (ver fluxograma):

  - Conta 205826 -> mantem dia e mes com 2 digitos e ano com 4 digitos
        ex: 27/07/2026 -> 27072026

  - Conta 200000 -> mantem dia, mes e ano com 2 digitos e adiciona "20" na frente
        ex: 27/07/2026 -> 20270726   (20 + DDMMYY)

  - Conta 205393 -> mantem dia, mes e ano com 2 digitos e adiciona "53" na frente
        ex: 27/07/2026 -> 53270726   (53 + DDMMYY)

Novo nome do arquivo: <data-formatada-pela-regra>.ret

Layout usado (CNAB240 Febraban):
  registro 0 (header de arquivo), colunas 59-70 (indice 58:70)
      -> numero da conta corrente (12 digitos, zero-padded)
  registro 3 (detalhe, segmento E), logo apos o literal "S"
      -> data de lancamento, formato DDMMAAAA (posicao 135-142, indice 134:142)

Uso:
    python3 renomear_ieddeb.py <arquivo_ou_pasta> [--dry-run]
"""

import argparse
import re
import sys
from pathlib import Path

# Regras: prefixo da conta encontrada no retorno -> (codigo usado no nome, funcao de formatacao)
REGRAS = {
    "205826": {
        "codigo": "205826",
        "formato": lambda d, m, y: f"{d:02d}{m:02d}{y:04d}",          # DDMMYYYY
    },
    "200000": {
        "codigo": "200000",
        "formato": lambda d, m, y: f"20{d:02d}{m:02d}{y % 100:02d}",  # 20 + DDMMYY
    },
    "205393": {
        "codigo": "205393",
        "formato": lambda d, m, y: f"53{d:02d}{m:02d}{y % 100:02d}",  # 53 + DDMMYY
    },
}

# Posicoes fixas (0-indexed, seguindo Python slicing) dentro do header de arquivo
CONTA_INICIO, CONTA_FIM = 58, 70
TAMANHO_MINIMO_HEADER = 172  # margem de seguranca; o registro tem 240 colunas

# Padrao do campo de data de lancamento dentro do registro de detalhe (segmento E):
# literal "S" seguido de duas datas DDMMAAAA (data de lancamento e data do balancete).
# Usamos a primeira, que e a data de lancamento.
DATA_LANCAMENTO_RE = re.compile(r"S(\d{2})(\d{2})(\d{4})\d{8}")


def extrair_conta(ret_path: Path) -> str:
    """Extrai o numero da conta a partir do header (primeira linha) do .ret."""
    with open(ret_path, encoding="latin-1") as f:
        header = f.readline().rstrip("\r\n")

    if len(header) < TAMANHO_MINIMO_HEADER:
        raise ValueError(
            f"Header muito curto ({len(header)} colunas) em {ret_path.name}; "
            "nao parece ser um retorno CNAB240 valido"
        )

    conta_bruta = header[CONTA_INICIO:CONTA_FIM].strip()
    return conta_bruta.lstrip("0") or "0"


def extrair_data_lancamento(ret_path: Path):
    """Extrai a data de lancamento a partir do primeiro registro de detalhe (segmento E)."""
    with open(ret_path, encoding="latin-1") as f:
        for linha in f:
            m = DATA_LANCAMENTO_RE.search(linha)
            if m:
                dia, mes, ano = (int(x) for x in m.groups())
                return dia, mes, ano

    raise ValueError(f"Nao foi possivel encontrar a data de lancamento em {ret_path.name}")


def extrair_dados(ret_path: Path):
    """Extrai o numero da conta e a data de lancamento do .ret."""
    conta = extrair_conta(ret_path)
    dia, mes, ano = extrair_data_lancamento(ret_path)
    return conta, dia, mes, ano


def identificar_regra(conta: str):
    """Encontra a regra aplicavel comparando o prefixo da conta encontrada."""
    for prefixo, regra in REGRAS.items():
        if conta.startswith(prefixo):
            return regra
    raise ValueError(f"Tipo de conta nao mapeado nas regras: {conta}")


def novo_nome(ret_path: Path) -> str:
    conta, dia, mes, ano = extrair_dados(ret_path)
    regra = identificar_regra(conta)
    data_formatada = regra["formato"](dia, mes, ano)
    return f"{data_formatada}{ret_path.suffix}"


def processar(caminho: Path, dry_run: bool = False):
    arquivos = [caminho] if caminho.is_file() else sorted(caminho.glob("IEDDEB*.ret"))

    if not arquivos:
        print(f"Nenhum arquivo IEDDEB*.ret encontrado em {caminho}")
        return

    for arquivo in arquivos:
        try:
            destino = arquivo.with_name(novo_nome(arquivo))
        except ValueError as e:
            print(f"[ERRO] {arquivo.name}: {e}")
            continue

        if destino == arquivo:
            print(f"[OK] {arquivo.name} ja esta no formato correto")
            continue

        if destino.exists():
            print(f"[ERRO] {arquivo.name}: destino {destino.name} ja existe, pulando")
            continue

        if dry_run:
            print(f"[DRY-RUN] {arquivo.name} -> {destino.name}")
        else:
            arquivo.rename(destino)
            print(f"[RENOMEADO] {arquivo.name} -> {destino.name}")


def main():
    parser = argparse.ArgumentParser(description="Renomeia arquivos IEDDEB (.ret) conforme o tipo de conta.")
    parser.add_argument("caminho", help="Arquivo .ret unico ou pasta contendo os arquivos IEDDEB*.ret")
    parser.add_argument("--dry-run", action="store_true", help="Mostra o que seria feito sem renomear de fato")
    args = parser.parse_args()

    caminho = Path(args.caminho)
    if not caminho.exists():
        print(f"Caminho nao encontrado: {caminho}")
        sys.exit(1)

    processar(caminho, dry_run=args.dry_run)


if __name__ == "__main__":
    main()