"""
Script para tratar a aba "Pagamentos" (gerada pelo separar_tipo.py).

Passo 1: remove da tabela de OBs (colunas "OB"/"VALOR") todas as linhas em
que VALOR = 0.

Passo 2: concilia pagamento(s) com OB. Para cada OB, procura um pagamento
(ou uma combinação de pagamentos ainda não usados) cuja soma bata
exatamente com o VALOR da OB. Quando acha, pinta a(s) linha(s) do(s)
pagamento(s) e a linha da OB com a mesma cor, deixando visualmente claro
quem está atrelado a quem.

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

# Cores usadas para marcar cada grupo (pagamento(s) <-> OB). Se houver mais
# grupos que cores, elas se repetem.
CORES_CONCILIACAO = [
    "FFF2CC", "D9EAD3", "CFE2F3", "F4CCCC", "EAD1DC",
    "D9D2E9", "FCE5CD", "D0E0E3", "B6D7A8", "F9CB9C",
]


def localizar_bloco_ob(ws):
    """Encontra as colunas 'OB' e 'VALOR' (lado a lado) pelo cabeçalho e
    devolve (col_ob, col_valor, linhas). (None, None, None) se não achar."""
    header_row = ws[1]
    col_ob = None
    for cell in header_row:
        if cell.value and str(cell.value).strip().upper() == "OB":
            col_ob = cell.column
            break
    if col_ob is None:
        return None, None, None

    col_valor = col_ob + 1
    valor_titulo = ws.cell(row=1, column=col_valor).value
    if not valor_titulo or str(valor_titulo).strip().upper() != "VALOR":
        return None, None, None

    linhas = []
    for r in range(2, ws.max_row + 1):
        v_ob = ws.cell(row=r, column=col_ob).value
        v_valor = ws.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            continue
        linhas.append((v_ob, v_valor))

    return col_ob, col_valor, linhas


def remover_valor_zero(caminho, aba="Pagamentos"):
    wb = openpyxl.load_workbook(caminho)
    if aba not in wb.sheetnames:
        raise ValueError(f"Aba '{aba}' não encontrada em {caminho}. Rode o separar_tipo.py primeiro.")
    ws = wb[aba]

    col_ob, col_valor, linhas = localizar_bloco_ob(ws)
    if col_ob is None:
        print(f"Bloco OB/VALOR não encontrado na aba '{aba}'. Nada foi alterado.")
        return

    total_antes = len(linhas)
    linhas_filtradas = [(ob, valor) for ob, valor in linhas if valor != 0]
    removidas = total_antes - len(linhas_filtradas)

    # limpa toda a área antiga do bloco (valor, borda e preenchimento)
    for r in range(2, ws.max_row + 1):
        for col in (col_ob, col_valor):
            cell = ws.cell(row=r, column=col)
            cell.value = None
            cell.border = NO_BORDER
            cell.fill = NO_FILL

    # reescreve só as linhas que sobraram
    for r, (ob, valor) in enumerate(linhas_filtradas, start=2):
        c_ob = ws.cell(row=r, column=col_ob, value=ob)
        c_ob.font = NORMAL_FONT
        c_ob.border = BORDER
        c_ob.alignment = CENTER

        c_valor = ws.cell(row=r, column=col_valor, value=valor)
        c_valor.font = NORMAL_FONT
        c_valor.border = BORDER
        c_valor.number_format = "#,##0.00"
        c_valor.alignment = RIGHT

    wb.save(caminho)

    print(f"OBs com VALOR = 0 removidas: {removidas}")
    print(f"OBs restantes na aba '{aba}': {len(linhas_filtradas)}")


# ---------------------------------------------------------------------------
# Passo 2: conciliação pagamento(s) <-> OB
# ---------------------------------------------------------------------------

def _centavos(valor):
    """Converte pra inteiro em centavos, pra comparar sem erro de ponto flutuante."""
    return round(float(valor) * 100)


def _linhas_pagamentos(ws, col_limite=7):
    """Lê a tabela principal (A:G) e devolve [(linha, quantia)] enquanto a
    coluna 'Linha' (A) tiver valor."""
    linhas = []
    r = 2
    while ws.cell(row=r, column=1).value is not None:
        quantia = ws.cell(row=r, column=5).value  # coluna E = Quantia
        if quantia is not None:
            linhas.append((r, quantia))
        r += 1
    return linhas


def _linhas_ob(ws, col_ob, col_valor):
    """Devolve [(linha, codigo_ob, valor)] do bloco de OBs."""
    linhas = []
    r = 2
    while True:
        v_ob = ws.cell(row=r, column=col_ob).value
        v_valor = ws.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            break
        if v_valor is not None:
            linhas.append((r, v_ob, v_valor))
        r += 1
    return linhas


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


def conciliar_pagamentos_obs(caminho, aba="Pagamentos", max_combinacao=6):
    wb = openpyxl.load_workbook(caminho)
    if aba not in wb.sheetnames:
        raise ValueError(f"Aba '{aba}' não encontrada em {caminho}. Rode o separar_tipo.py primeiro.")
    ws = wb[aba]

    col_ob, col_valor, _ = localizar_bloco_ob(ws)
    if col_ob is None:
        print(f"Bloco OB/VALOR não encontrado na aba '{aba}'. Nada foi feito.")
        return

    pagamentos = _linhas_pagamentos(ws)
    obs = _linhas_ob(ws, col_ob, col_valor)

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

    wb.save(caminho)

    print(f"OBs conciliadas: {len(grupos)} de {len(obs)}")
    for grupo in grupos:
        n_pag = len(grupo["pagamento_rows"])
        composicao = "1 pagamento" if n_pag == 1 else f"{n_pag} pagamentos somados"
        print(f"  OB {grupo['ob_codigo']} (R$ {grupo['valor']:.2f}) <- {composicao}, linha(s) {grupo['pagamento_rows']}")

    if obs_sem_match:
        print(f"OBs SEM correspondência encontrada: {len(obs_sem_match)}")
        for ob_row, ob_codigo, valor in obs_sem_match:
            print(f"  OB {ob_codigo} (R$ {valor:.2f})")

    if disponiveis:
        print(f"Pagamentos sem OB correspondente: {len(disponiveis)}")
        for linha, quantia in disponiveis:
            print(f"  Linha {linha} (R$ {quantia:.2f})")


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "conciliacao.xlsx"
    remover_valor_zero(caminho)
    print()
    conciliar_pagamentos_obs(caminho)