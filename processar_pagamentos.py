"""
Script para tratar as abas geradas pelo separar_tipo.py (uma aba por dia,
cada uma com um bloco "PAGAMENTOS" e, ao lado, um bloco "OB"/"VALOR").

Passo 1: em cada aba, remove do bloco OB/VALOR todas as linhas em que
VALOR = 0.

Passo 2: em cada aba, concilia pagamento(s) com OB. Para cada OB, procura
um pagamento (ou uma combinação de pagamentos ainda não usados, daquela
mesma aba) cuja soma bata exatamente com o VALOR da OB. Quando acha, pinta
a(s) linha(s) do(s) pagamento(s) e a linha da OB com a mesma cor.

Como o separar_tipo.py processa cada aba de forma independente e a posição
das linhas varia (depende de quantos Recebimentos existem antes), este
script localiza os blocos dinamicamente em cada aba, em vez de assumir uma
aba fixa chamada "Pagamentos" com cabeçalho na linha 1.

Uso:
    python processar_pagamentos.py caminho_da_planilha.xlsx
"""

import sys
from itertools import combinations

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
NORMAL_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NO_BORDER = Border()
NO_FILL = PatternFill(fill_type=None)
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

COL_QUANTIA = 5  # coluna E da tabela principal (A:G), igual em toda aba

# Cores usadas para marcar cada grupo (pagamento(s) <-> OB). Se houver mais
# grupos que cores, elas se repetem.
CORES_CONCILIACAO = [
    "FFF2CC", "D9EAD3", "CFE2F3", "F4CCCC", "EAD1DC",
    "D9D2E9", "FCE5CD", "D0E0E3", "B6D7A8", "F9CB9C",
]


# ---------------------------------------------------------------------------
# Localização dos blocos dentro de uma aba já processada pelo separar_tipo.py
# ---------------------------------------------------------------------------

def localizar_bloco_pagamentos(ws):
    """Acha o título 'PAGAMENTOS (...)' (coluna A) na aba e devolve
    (linha_header, linha_inicio_dados), ou (None, None) se não achar."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).strip().upper().startswith("PAGAMENTOS"):
            return r + 1, r + 2
    return None, None


def localizar_bloco_ob(ws):
    """Procura em toda a aba um cabeçalho 'OB' seguido, na coluna ao lado,
    por 'VALOR'. Devolve (col_ob, col_valor, linha_inicio_dados) ou
    (None, None, None)."""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v and str(v).strip().upper() == "OB":
                v2 = ws.cell(row=r, column=c + 1).value
                if v2 and str(v2).strip().upper() == "VALOR":
                    return c, c + 1, r + 1
    return None, None, None


def _linhas_ob(ws, col_ob, col_valor, linha_inicio):
    """Devolve [(linha, codigo_ob, valor)] a partir de linha_inicio,
    enquanto houver algo em OB ou VALOR naquela linha."""
    linhas = []
    r = linha_inicio
    while True:
        v_ob = ws.cell(row=r, column=col_ob).value
        v_valor = ws.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            break
        linhas.append((r, v_ob, v_valor))
        r += 1
    return linhas


def _linhas_pagamentos(ws, linha_inicio):
    """Lê a tabela principal (A:G) a partir de linha_inicio, enquanto a
    coluna 'Linha' (A) tiver valor. Devolve [(linha, quantia)]."""
    linhas = []
    r = linha_inicio
    while ws.cell(row=r, column=1).value is not None:
        quantia = ws.cell(row=r, column=COL_QUANTIA).value
        if quantia is not None:
            linhas.append((r, quantia))
        r += 1
    return linhas


# ---------------------------------------------------------------------------
# Passo 1: remove OBs com VALOR = 0 (em cada aba)
# ---------------------------------------------------------------------------

def remover_valor_zero(caminho):
    wb = openpyxl.load_workbook(caminho)

    total_removidas = 0
    total_restantes = 0
    abas_sem_bloco = []

    for nome in wb.sheetnames:
        ws = wb[nome]
        col_ob, col_valor, linha_inicio = localizar_bloco_ob(ws)
        if col_ob is None:
            abas_sem_bloco.append(nome)
            continue

        linhas = _linhas_ob(ws, col_ob, col_valor, linha_inicio)
        linhas_filtradas = [(ob, valor) for _, ob, valor in linhas if valor != 0]
        removidas = len(linhas) - len(linhas_filtradas)

        # limpa toda a área antiga do bloco (valor, borda, preenchimento) -
        # exatamente a faixa que continha dados antes da filtragem
        for r in range(linha_inicio, linha_inicio + len(linhas)):
            for col in (col_ob, col_valor):
                cell = ws.cell(row=r, column=col)
                cell.value = None
                cell.border = NO_BORDER
                cell.fill = NO_FILL

        # reescreve só as linhas que sobraram
        for r, (ob, valor) in enumerate(linhas_filtradas, start=linha_inicio):
            c_ob = ws.cell(row=r, column=col_ob, value=ob)
            c_ob.font = NORMAL_FONT
            c_ob.border = BORDER
            c_ob.alignment = CENTER

            c_valor = ws.cell(row=r, column=col_valor, value=valor)
            c_valor.font = NORMAL_FONT
            c_valor.border = BORDER
            c_valor.number_format = "#,##0.00"
            c_valor.alignment = RIGHT

        total_removidas += removidas
        total_restantes += len(linhas_filtradas)
        print(f"[{nome}] OBs com VALOR = 0 removidas: {removidas} | restantes: {len(linhas_filtradas)}")

    wb.save(caminho)

    print()
    print(f"Total de OBs removidas: {total_removidas}")
    print(f"Total de OBs restantes: {total_restantes}")
    if abas_sem_bloco:
        print(f"Abas sem bloco OB/VALOR (nada feito nelas): {len(abas_sem_bloco)}")
        print("  ->", ", ".join(abas_sem_bloco))


# ---------------------------------------------------------------------------
# Passo 2: conciliação pagamento(s) <-> OB (em cada aba)
# ---------------------------------------------------------------------------

def _centavos(valor):
    """Converte pra inteiro em centavos, pra comparar sem erro de ponto flutuante."""
    return round(float(valor) * 100)


def _eh_numero(v):
    """True se v é um número usável como valor monetário (exclui bool,
    texto, None etc.)."""
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _buscar_subconjunto(pagamentos_disponiveis, alvo_centavos, max_tamanho):
    """Procura, do menor pro maior tamanho, uma combinação de pagamentos
    (linha, quantia) cuja soma em centavos bate exatamente com o alvo.
    Devolve a combinação (tupla de (linha, quantia)) ou None."""
    n = len(pagamentos_disponiveis)
    max_tamanho = min(max_tamanho, n)
    for tamanho in range(1, max_tamanho + 1):
        for combo in combinations(pagamentos_disponiveis, tamanho):
            if sum(_centavos(q) for _, q in combo) == alvo_centavos:
                return combo
    return None


def conciliar_pagamentos_obs(caminho, max_combinacao=6):
    wb = openpyxl.load_workbook(caminho)

    abas_processadas = 0
    abas_sem_bloco = []

    for nome in wb.sheetnames:
        ws = wb[nome]

        linha_header_pag, linha_dados_pag = localizar_bloco_pagamentos(ws)
        col_ob, col_valor, linha_dados_ob = localizar_bloco_ob(ws)

        if linha_header_pag is None or col_ob is None:
            abas_sem_bloco.append(nome)
            continue

        pagamentos = _linhas_pagamentos(ws, linha_dados_pag)
        obs_todas = _linhas_ob(ws, col_ob, col_valor, linha_dados_ob)

        # Algumas abas têm, coladas nas mesmas colunas logo abaixo da lista
        # de OB/VALOR, outras informações que nada têm a ver com a
        # conciliação (ex.: nome de responsável, categoria de despesa).
        # Só entram na conciliação as linhas com VALOR realmente numérico;
        # o resto é reportado à parte, sem ser tocado.
        obs = [(r, ob, valor) for r, ob, valor in obs_todas if _eh_numero(valor)]
        obs_fora_padrao = [(r, ob, valor) for r, ob, valor in obs_todas if not _eh_numero(valor)]

        # OBs maiores primeiro: tende a reduzir ambiguidade na hora de casar combinações
        obs_ordenadas = sorted(obs, key=lambda x: x[2], reverse=True)

        disponiveis = list(pagamentos)  # (linha, quantia) ainda não atrelados a nenhuma OB
        grupos = []
        obs_sem_match = []

        for ob_row, ob_codigo, valor in obs_ordenadas:
            alvo = _centavos(valor)
            combo = _buscar_subconjunto(disponiveis, alvo, max_combinacao)
            if combo is None:
                obs_sem_match.append((ob_row, ob_codigo, valor))
                continue
            linhas_pagamento = [linha for linha, _ in combo]
            grupos.append({
                "ob_row": ob_row,
                "ob_codigo": ob_codigo,
                "valor": valor,
                "pagamento_rows": linhas_pagamento,
            })
            usados = set(linhas_pagamento)
            disponiveis = [p for p in disponiveis if p[0] not in usados]

        # pinta os grupos conciliados
        for i, grupo in enumerate(grupos):
            cor = CORES_CONCILIACAO[i % len(CORES_CONCILIACAO)]
            fill = PatternFill("solid", fgColor=cor)

            for col in (col_ob, col_valor):
                ws.cell(row=grupo["ob_row"], column=col).fill = fill

            for linha in grupo["pagamento_rows"]:
                for col in range(1, 8):  # A:G
                    ws.cell(row=linha, column=col).fill = fill

        abas_processadas += 1
        print(f"[{nome}] OBs conciliadas: {len(grupos)} de {len(obs)}")
        for grupo in grupos:
            n_pag = len(grupo["pagamento_rows"])
            composicao = "1 pagamento" if n_pag == 1 else f"{n_pag} pagamentos somados"
            print(f"    OB {grupo['ob_codigo']} (R$ {grupo['valor']:.2f}) <- {composicao}, linha(s) {grupo['pagamento_rows']}")
        if obs_sem_match:
            print(f"    OBs SEM correspondência encontrada: {len(obs_sem_match)}")
            for ob_row, ob_codigo, valor in obs_sem_match:
                print(f"      OB {ob_codigo} (R$ {valor:.2f})")
        if disponiveis:
            print(f"    Pagamentos sem OB correspondente: {len(disponiveis)}")
            for linha, quantia in disponiveis:
                print(f"      Linha {linha} (R$ {quantia:.2f})")
        if obs_fora_padrao:
            print(f"    ATENÇÃO: {len(obs_fora_padrao)} linha(s) nas colunas OB/VALOR com "
                  f"conteúdo não numérico (ignoradas na conciliação, nada foi alterado nelas):")
            for r, ob, valor in obs_fora_padrao[:10]:
                print(f"      - linha {r}: OB={ob!r} VALOR={valor!r}")
            if len(obs_fora_padrao) > 10:
                print(f"      ... e mais {len(obs_fora_padrao) - 10}")

    wb.save(caminho)

    print()
    print(f"Abas processadas: {abas_processadas}")
    if abas_sem_bloco:
        print(f"Abas sem bloco PAGAMENTOS/OB (nada feito nelas): {len(abas_sem_bloco)}")
        print("  ->", ", ".join(abas_sem_bloco))


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "conciliacao.xlsx"
    remover_valor_zero(caminho)
    print()
    conciliar_pagamentos_obs(caminho)