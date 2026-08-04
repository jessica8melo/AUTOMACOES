"""
Script para separar a planilha de conciliação em Recebimentos e Pagamentos.

Regra usada: qualquer valor da coluna "Tipo" que COMEÇA com "Recebimento"
vai para a aba/lista de Recebimentos; qualquer valor que COMEÇA com
"Pagamento" vai para a aba/lista de Pagamentos. Isso cobre variações como
"Recebimento" e "Recebimento Div." / "Pagamento" e "Pagamento Div."

Uso:
    python separar_tipo.py caminho_da_planilha.xlsx
"""

import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
NORMAL_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")


def classificar(tipo: str) -> str:
    """Retorna 'Recebimento', 'Pagamento' ou 'Outro' com base no texto do Tipo."""
    if not tipo:
        return "Outro"
    t = str(tipo).strip().lower()
    if t.startswith("recebimento"):
        return "Recebimento"
    if t.startswith("pagamento"):
        return "Pagamento"
    return "Outro"


def escrever_aba(wb, nome_aba, headers, linhas, col_valor_idx=None):
    """Cria (ou recria) uma aba, escreve o cabeçalho estilizado e as linhas."""
    if nome_aba in wb.sheetnames:
        del wb[nome_aba]
    ws = wb.create_sheet(nome_aba)

    for col, titulo in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=titulo)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = BORDER

    for r, linha in enumerate(linhas, start=2):
        for c_idx, valor in enumerate(linha, start=1):
            cell = ws.cell(row=r, column=c_idx, value=valor)
            cell.font = NORMAL_FONT
            cell.border = BORDER
            if col_valor_idx is not None and c_idx == col_valor_idx:
                cell.number_format = "#,##0.00"
                cell.alignment = RIGHT
            elif c_idx in (1, 4):  # Linha / Data Transação centralizadas
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT

    if linhas:
        ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(headers)).coordinate[0]}{len(linhas) + 1}"

    larguras = {"A": 7, "B": 18, "C": 11, "D": 15, "E": 13, "F": 18, "G": 32}
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    return ws


def main(caminho):
    wb = openpyxl.load_workbook(caminho)
    ws_origem = wb["Conciliação"] if "Conciliação" in wb.sheetnames else wb.active

    headers = [c.value for c in ws_origem[1][:7]]  # A:G

    recebimentos = []
    pagamentos = []
    outros = []

    for row in ws_origem.iter_rows(min_row=2, max_col=7, values_only=True):
        if row[0] is None:
            continue
        tipo = row[1]
        categoria = classificar(tipo)
        if categoria == "Recebimento":
            recebimentos.append(row)
        elif categoria == "Pagamento":
            pagamentos.append(row)
        else:
            outros.append(row)

    escrever_aba(wb, "Recebimentos", headers, recebimentos, col_valor_idx=5)
    escrever_aba(wb, "Pagamentos", headers, pagamentos, col_valor_idx=5)
    if outros:
        escrever_aba(wb, "Outros", headers, outros, col_valor_idx=5)

    wb.save(caminho)

    print(f"Recebimentos: {len(recebimentos)} linhas")
    print(f"Pagamentos:   {len(pagamentos)} linhas")
    if outros:
        print(f"Outros:       {len(outros)} linhas (não classificados)")


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "conciliacao.xlsx"
    main(caminho)
