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


def localizar_bloco_ob(ws_origem):
    """Procura, na linha de cabeçalho, colunas chamadas 'OB' e 'VALOR' (nessa
    ordem, lado a lado) e devolve (headers, linhas) desse bloco, ou (None, None)
    se não encontrar. Independe da posição exata das colunas na planilha."""
    header_row = ws_origem[1]
    col_ob = None
    for cell in header_row:
        if cell.value and str(cell.value).strip().upper() == "OB":
            col_ob = cell.column
            break
    if col_ob is None:
        return None, None

    col_valor = col_ob + 1
    valor_titulo = ws_origem.cell(row=1, column=col_valor).value
    if not valor_titulo or str(valor_titulo).strip().upper() != "VALOR":
        return None, None

    headers = [ws_origem.cell(row=1, column=col_ob).value,
               ws_origem.cell(row=1, column=col_valor).value]

    linhas = []
    for r in range(2, ws_origem.max_row + 1):
        v_ob = ws_origem.cell(row=r, column=col_ob).value
        v_valor = ws_origem.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            continue
        linhas.append((v_ob, v_valor))

    return headers, linhas


def escrever_aba(wb, nome_aba, headers, linhas, col_valor_idx=None, bloco_extra=None):
    """Cria (ou recria) uma aba, escreve o cabeçalho estilizado e as linhas.

    bloco_extra (opcional): tupla (headers_extra, linhas_extra, coluna_inicial)
    escrita ao lado da tabela principal, com uma coluna de espaço entre elas.
    """
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

    if bloco_extra is not None:
        headers_extra, linhas_extra, col_inicial = bloco_extra
        for col_off, titulo in enumerate(headers_extra):
            c = ws.cell(row=1, column=col_inicial + col_off, value=titulo)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.alignment = CENTER
            c.border = BORDER

        for r, linha in enumerate(linhas_extra, start=2):
            for col_off, valor in enumerate(linha):
                cell = ws.cell(row=r, column=col_inicial + col_off, value=valor)
                cell.font = NORMAL_FONT
                cell.border = BORDER
                if col_off == 1:  # coluna VALOR
                    cell.number_format = "#,##0.00"
                    cell.alignment = RIGHT
                else:
                    cell.alignment = CENTER

        letra_inicial = ws.cell(row=1, column=col_inicial).column_letter
        ws.column_dimensions[letra_inicial].width = 14
        ws.column_dimensions[ws.cell(row=1, column=col_inicial + 1).column_letter].width = 14

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

    headers_ob, linhas_ob = localizar_bloco_ob(ws_origem)
    bloco_ob = None
    if headers_ob is not None:
        # deixa uma coluna de espaço (H) após a tabela principal (A:G)
        col_inicial_ob = len(headers) + 2  # 7 colunas + 1 de espaço + 1 = 9 (col I)
        bloco_ob = (headers_ob, linhas_ob, col_inicial_ob)

    escrever_aba(wb, "Recebimentos", headers, recebimentos, col_valor_idx=5)
    escrever_aba(wb, "Pagamentos", headers, pagamentos, col_valor_idx=5, bloco_extra=bloco_ob)
    if outros:
        escrever_aba(wb, "Outros", headers, outros, col_valor_idx=5)

    wb.save(caminho)

    print(f"Recebimentos: {len(recebimentos)} linhas")
    print(f"Pagamentos:   {len(pagamentos)} linhas")
    if outros:
        print(f"Outros:       {len(outros)} linhas (não classificados)")
    if headers_ob is not None:
        print(f"Bloco OB/VALOR: {len(linhas_ob)} linhas levadas para a aba Pagamentos")
    else:
        print("Bloco OB/VALOR: não encontrado na planilha de origem (colunas 'OB'/'VALOR' não localizadas)")


if __name__ == "__main__":
    caminho = sys.argv[1] if len(sys.argv) > 1 else "conciliacao.xlsx"
    main(caminho)