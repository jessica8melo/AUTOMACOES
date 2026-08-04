"""
Script para tratar a aba "Pagamentos" (gerada pelo separar_tipo.py).

Passo 1: remove da tabela de OBs (colunas "OB"/"VALOR") todas as linhas em
que VALOR = 0.

Uso:
    python processar_pagamentos.py caminho_da_planilha.xlsx
"""

import sys
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


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "conciliacao.xlsx"
    remover_valor_zero(caminho)
