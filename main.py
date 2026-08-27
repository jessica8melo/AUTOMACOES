# -*- coding: utf-8 -*-
"""
Script principal: encadeia a análise do controle financeiro com o
cruzamento no relatório de pagamentos.

Fluxo:
    1. Roda controle_financeiro.analisar() sobre a planilha de controle
       financeiro, obtendo todas as ocorrências com a célula "OB" pintada
       de amarelo e vazia (cada ocorrência já vem com Data, Descrição,
       Total OB, Código, CAT e ORG).
    2. Para cada CAT distinto encontrado nessas ocorrências, roda
       relatorio_pagamentos.filtrar_por_cat() sobre a planilha de
       pagamentos, trazendo as linhas de "Fornecedor" que batem com aquele
       CAT, e soma os valores por "Data Pagamento".

Uso:
    python main.py [--controle CAMINHO.xlsx] [--sheet-controle NOME]
                    [--pagamento CAMINHO.xlsx] [--sheet-pagamento NOME]

Se os caminhos não forem informados, usam os valores padrão definidos nos
próprios módulos (DEFAULT_PATH de cada um).

Requisito: controle_financeiro.py, relatorio_pagamentos.py e
tabela_localidade.py devem estar na mesma pasta deste script (ou no
PYTHONPATH).
"""

import argparse

from controle_financeiro import analisar as analisar_controle
from controle_financeiro import DEFAULT_PATH as DEFAULT_CONTROLE
from relatorio_pagamentos import filtrar_por_cat, agrupar_valor_por_data
from relatorio_pagamentos import DEFAULT_PATH as DEFAULT_PAGAMENTO


def imprimir_ocorrencias_controle(ocorrencias):
    print("=" * 70)
    print("ETAPA 1 — Controle financeiro (células OB amarelas e vazias)")
    print("=" * 70)

    if not ocorrencias:
        print("Nenhuma ocorrência com célula OB amarela foi encontrada.\n")
        return

    print(f"Total de ocorrências encontradas: {len(ocorrencias)}\n")
    for i, item in enumerate(ocorrencias, start=1):
        cat = item["CAT"] or "não encontrado"
        org = item["ORG"] or "não encontrado"
        print(
            f"{i}. Data: {item['Data']} | Descrição: {item['Descrição']} | "
            f"Total OB: {item['Total OB']} | Código: {item['Código']} | "
            f"CAT: {cat} | ORG: {org}"
        )
    print()


def processar_pagamentos_por_cat(caminho_pagamento, cats, sheet_pagamento=None):
    print("=" * 70)
    print("ETAPA 2 — Relatório de pagamentos (cruzado por CAT)")
    print("=" * 70)

    if not cats:
        print("Nenhum CAT válido foi identificado na etapa 1; nada a cruzar.\n")
        return

    for cat in cats:
        print(f"\n--- CAT: {cat} ---")
        try:
            registro_cat, resultados = filtrar_por_cat(caminho_pagamento, cat, sheet_pagamento)
        except ValueError as erro:
            print(f"Aviso: {erro}")
            continue

        print(
            f"CAT: {registro_cat['cat']}  |  Sigla: {registro_cat['sigla']}  |  "
            f"ORG: {registro_cat['org']}"
        )

        if not resultados:
            print("Nenhuma linha de Fornecedor encontrada para esse CAT.")
            continue

        print(f"Total de ocorrências encontradas: {len(resultados)}\n")
        for i, item in enumerate(resultados, start=1):
            print(
                f"{i}. Fornecedor: {item['Fornecedor']} | Grupo: {item['Grupo Pagamentos']} | "
                f"NFF: {item['NFF']} | Data Pagamento: {item['Data Pagamento']} | "
                f"Valor: {item['Valor']}"
            )

        print("\nSoma de Valor por Data Pagamento:")
        totais = agrupar_valor_por_data(resultados)
        for grupo in totais:
            data = grupo["Data Pagamento"] or "(sem data)"
            print(f"{data}: {grupo['Total']:.2f}")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Roda a análise do controle financeiro e, em seguida, cruza cada "
            "CAT encontrado com o relatório de pagamentos."
        )
    )
    parser.add_argument(
        "--controle", default=DEFAULT_CONTROLE,
        help="Caminho do arquivo .xlsx do controle financeiro"
    )
    parser.add_argument(
        "--sheet-controle", default=None,
        help="Nome da aba do controle financeiro (padrão: aba ativa)"
    )
    parser.add_argument(
        "--pagamento", default=DEFAULT_PAGAMENTO,
        help="Caminho do arquivo .xlsx do relatório de pagamento"
    )
    parser.add_argument(
        "--sheet-pagamento", default=None,
        help="Nome da aba do relatório de pagamento (padrão: aba ativa)"
    )
    args = parser.parse_args()

    # Etapa 1: controle_financeiro
    ocorrencias = analisar_controle(args.controle, args.sheet_controle)
    imprimir_ocorrencias_controle(ocorrencias)

    # Coleta os CATs distintos (na ordem em que aparecem, sem repetir)
    cats_distintos = []
    for item in ocorrencias:
        cat = item["CAT"]
        if cat and cat not in cats_distintos:
            cats_distintos.append(cat)

    # Etapa 2: relatorio_pagamentos, um cruzamento por CAT distinto
    processar_pagamentos_por_cat(args.pagamento, cats_distintos, args.sheet_pagamento)


if __name__ == "__main__":
    main()