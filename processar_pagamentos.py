"""
Script para tratar as abas geradas pelo separar_tipo.py (uma aba por dia,
cada uma já dividida em blocos RECEBIMENTOS / PAGAMENTOS, com uma tabela
auxiliar "OB" / "VALOR" à direita, dentro do bloco de Pagamentos).

Percorre TODAS as abas da planilha e trata cada uma de forma independente,
já que cada dia tem seu próprio bloco de Pagamentos e sua própria tabela de
OBs.

A lógica está dividida em módulos, cada um cuidando de uma etapa:

    - estilos.py           constantes de fonte/cor/borda
    - helpers_planilha.py  helpers de localização/leitura genéricos
    - remover_zero.py      Passo 1: remove OBs com VALOR = 0 (rodar antes,
                            separadamente - ver README)
    - conciliacao.py       Passo 2: concilia pagamento(s) <-> OB
    - extratos.py          Passo 3 (opcional): cruza com extratos bancários

Este script cuida só dos Passos 2 e 3 (conciliação + Etapa 3 opcional). O
Passo 1 (remover_zero) não é mais chamado automaticamente aqui - rode-o à
parte antes, com `python remover_zero.py caminho_da_planilha.xlsx`.

Ver a docstring de cada módulo para o detalhe de cada passo.

Uso:
    python processar_pagamentos.py caminho_da_planilha.xlsx
    python processar_pagamentos.py caminho_da_planilha.xlsx --extratos saida.xlsx
"""

import argparse

from conciliacao import conciliar_pagamentos_obs
from extratos import atualizar_status_sem_conciliacao


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Processa pagamentos e OBs em uma planilha Excel.")
    parser.add_argument("caminho", nargs="?", default="conciliacao.xlsx", help="Caminho da planilha Excel")
    parser.add_argument("--aba", "--abas", dest="abas", action="append", help="Nome da aba a ser processada. Pode ser informado mais de uma vez.")
    parser.add_argument("--extratos", dest="extratos", default=None, help="Caminho da planilha de extratos (saida.xlsx). Se informado, roda também o passo 3 (marcar TARIFA / status nos pagamentos sem OB).")
    parser.add_argument("--max-combinacao-pagamentos", dest="max_combinacao", type=int, default=11, help="Tamanho máximo de combinação de pagamentos (2 ou mais) testada nas prioridades 3 e 4, quando uma OB isolada não bate com nenhum pagamento sozinho. O escalonamento interno começa em 6 e só sobe até este valor se sobrarem OBs sem match. Padrão: 11.")
    args = parser.parse_args()

    if args.extratos:
        # já conciliamos por OB dentro de atualizar_status_sem_conciliacao,
        # então não precisa chamar conciliar_pagamentos_obs de novo aqui.
        atualizar_status_sem_conciliacao(
            args.caminho, args.extratos, max_combinacao=args.max_combinacao, abas=args.abas
        )
    else:
        conciliar_pagamentos_obs(args.caminho, max_combinacao=args.max_combinacao, abas=args.abas)