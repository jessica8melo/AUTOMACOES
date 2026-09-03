# -*- coding: utf-8 -*-
"""
Script principal: encadeia a análise do controle financeiro, o
cruzamento no relatório de pagamentos e a conferência no relatório de
recebimento integrado.

Fluxo:
    1. Roda controle_financeiro.analisar() sobre a planilha de controle
       financeiro, obtendo todas as ocorrências com a célula "OB" pintada
       de amarelo e vazia (cada ocorrência já vem com Data, Descrição,
       Total OB, Código, CAT e ORG).
    2. Para cada CAT distinto encontrado nessas ocorrências, roda
       relatorio_pagamentos.filtrar_por_cat() sobre a planilha de
       pagamentos, trazendo as linhas de "Fornecedor" que batem com aquele
       CAT, e soma os valores por "Data Pagamento". Os pagamentos
       encontrados são agrupados pela ORG do respectivo CAT.
    3. Para cada ORG com pagamentos da etapa 2, roda
       relatorio_recebimento.verificar_por_org() sobre a planilha de
       Recebimento Integrado, conferindo N.F == NFF e ISS == Valor; quando
       ambos batem, pinta a célula "N.F" de amarelo.

Uso:
    python main.py [--controle CAMINHO.xlsx] [--sheet-controle NOME]
                    [--pagamento CAMINHO.xlsx] [--sheet-pagamento NOME]
                    [--recebimento CAMINHO.xlsx] [--sheet-recebimento NOME]
                    [--nao-salvar-recebimento]
                    [--cat "Nome do CAT"]

Se os caminhos não forem informados, usam os valores padrão definidos nos
próprios módulos (DEFAULT_PATH de cada um).

Rodar apenas um CAT específico:
    Use --cat "Nome do CAT" para pular a Etapa 1 (controle financeiro) e
    rodar as Etapas 2 e 3 apenas para o CAT informado, ex:

        python main.py --cat "Bauru"

    Nesse modo o script não varre o controle financeiro em busca de todos
    os CATs pendentes; ele cruza diretamente esse CAT com o relatório de
    pagamentos (Etapa 2) e, em seguida, com o relatório de recebimento
    integrado da ORG correspondente (Etapa 3).

Requisito: controle_financeiro.py, relatorio_pagamentos.py,
relatorio_recebimento.py e tabela_localidade.py devem estar na mesma pasta
deste script (ou no PYTHONPATH).
"""

import argparse

from controle_financeiro import analisar as analisar_controle
from controle_financeiro import DEFAULT_PATH as DEFAULT_CONTROLE
from relatorio_pagamentos import filtrar_por_cat, agrupar_valor_por_data
from relatorio_pagamentos import DEFAULT_PATH as DEFAULT_PAGAMENTO
from relatorio_recebimento import verificar_por_org, imprimir_resultados as imprimir_resultados_recebimento
from relatorio_recebimento import DEFAULT_PATH as DEFAULT_RECEBIMENTO


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

    pagamentos_por_org = {}

    if not cats:
        print("Nenhum CAT válido foi identificado na etapa 1; nada a cruzar.\n")
        return pagamentos_por_org

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

        # Acumula os pagamentos encontrados sob a ORG desse CAT, para a
        # etapa 3 (relatório de recebimento é cruzado por ORG, não por CAT).
        org = registro_cat.get("org")
        if org:
            pagamentos_por_org.setdefault(org, []).extend(resultados)

    return pagamentos_por_org


def processar_recebimento_por_org(caminho_recebimento, pagamentos_por_org, sheet_recebimento=None, salvar=True):
    print("=" * 70)
    print("ETAPA 3 — Relatório de Recebimento Integrado (cruzado por ORG)")
    print("=" * 70)

    if not pagamentos_por_org:
        print("Nenhum pagamento disponível da etapa 2; nada a cruzar.\n")
        return

    for org, pagamentos in pagamentos_por_org.items():
        try:
            resultados = verificar_por_org(
                caminho_recebimento, org, pagamentos,
                sheet=sheet_recebimento, salvar=salvar,
            )
        except ValueError as erro:
            print(f"Aviso: {erro}")
            continue

        imprimir_resultados_recebimento(org, resultados)


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
    parser.add_argument(
        "--recebimento", default=DEFAULT_RECEBIMENTO,
        help="Caminho do arquivo .xlsx do Relatório de Recebimento Integrado"
    )
    parser.add_argument(
        "--sheet-recebimento", default=None,
        help="Nome da aba do recebimento (padrão: aba ativa)"
    )
    parser.add_argument(
        "--nao-salvar-recebimento", action="store_true",
        help="Não grava a cor amarela na coluna N.F do recebimento (só mostra o resultado)"
    )
    parser.add_argument(
        "--cat", default=None,
        help=(
            'Roda apenas para este CAT específico (ex: "Bauru"), pulando a '
            "Etapa 1 (controle financeiro) e indo direto para as Etapas 2 e 3."
        ),
    )
    args = parser.parse_args()

    if args.cat:
        # Modo "um CAT específico": pula a Etapa 1 por completo.
        cats_distintos = [args.cat]
    else:
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
    pagamentos_por_org = processar_pagamentos_por_cat(
        args.pagamento, cats_distintos, args.sheet_pagamento
    )

    # Etapa 3: relatorio_recebimento, um cruzamento por ORG
    processar_recebimento_por_org(
        args.recebimento, pagamentos_por_org,
        sheet_recebimento=args.sheet_recebimento,
        salvar=not args.nao_salvar_recebimento,
    )


if __name__ == "__main__":
    main()