"""
Constantes de estilo (fonte, cores, bordas) usadas ao escrever/pintar
células nas planilhas, compartilhadas pelos outros módulos.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NO_BORDER = Border()
NO_FILL = PatternFill(fill_type=None)
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

NUM_COLS_PRINCIPAL = 7  # A:G -> Linha, Tipo, Código, Data Transação, Quantia, Status, Justificativas

# Cores usadas para marcar cada grupo (pagamento(s) <-> OB). Se houver mais
# grupos que cores, elas se repetem.
CORES_CONCILIACAO = [
    "FFD966",  # amarelo forte
    "93C47D",  # verde
    "6FA8DC",  # azul
    "E06666",  # vermelho
    "C27BA0",  # rosa/magenta
    "8E7CC3",  # roxo
    "F6B26B",  # laranja
    "76A5AF",  # verde-azulado (teal)
    "A2C4C9",  # azul petróleo claro
    "FF9999",  # salmão
    "B4A7D6",  # lilás
    "FFE599",  # amarelo claro (variação, ainda contrastante)
    "45818E",  # teal escuro
    "CC4125",  # vermelho tijolo
    "674EA7",  # roxo escuro
    "38761D",  # verde escuro
    "0B5394",  # azul escuro
    "BF9000",  # dourado/mostarda
    "D5A6BD",  # rosa antigo
    "999999",  # cinza (reserva pra quando esgotar as outras)
    "6D9EEB",  # azul claro
    "8FCE00",  # verde limão
    "E69138",  # laranja queimado
    "A64D79",  # vinho/magenta escuro
    "16537E",  # azul petróleo escuro
    "B45F06",  # marrom alaranjado
    "783F04",  # marrom escuro
    "76D7C4",  # verde-água claro
    "F9CB9C",  # pêssego claro
    "D9D2E9",  # lilás muito claro
    "D0E0E3",  # azul gelo
    "EAD1DC",  # rosa bebê
    "FCE5CD",  # creme alaranjado
    "C9DAF8",  # azul bebê
    "D9EAD3",  # verde bem claro
    "FFF2CC",  # amarelo bem claro
    "B6D7A8",  # verde pastel
    "A4C2F4",  # azul pastel
    "F4CCCC",  # vermelho bem claro
]