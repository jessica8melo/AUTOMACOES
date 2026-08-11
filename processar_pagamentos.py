"""
Script para tratar as abas geradas pelo separar_tipo.py (uma aba por dia,
cada uma já dividida em blocos RECEBIMENTOS / PAGAMENTOS, com uma tabela
auxiliar "OB" / "VALOR" à direita, dentro do bloco de Pagamentos).

Diferente da versão anterior (que esperava uma única aba fixa chamada
"Pagamentos"), esta percorre TODAS as abas da planilha e trata cada uma de
forma independente, já que cada dia tem seu próprio bloco de Pagamentos e
sua própria tabela de OBs.

Passo 1: em cada aba, remove da tabela de OBs (colunas "OB"/"VALOR") todas
as linhas em que VALOR = 0.

Passo 2: em cada aba, concilia pagamento(s) com OB dentro dessa mesma aba.
Para cada OB, a busca por pagamento(s) segue uma ordem de prioridade:

    1) Match exato de UM único pagamento cujo valor bate com o da OB.
       Se existir, é usado na hora, sem procurar combinações.
    2) Combinações de 2 ou mais pagamentos cuja soma bate com o valor da
       OB, dando preferência para pagamentos que estejam fisicamente
       próximos entre si na planilha (idealmente na linha logo abaixo ou
       logo acima um do outro).
    3) Se nada acima servir, vale qualquer combinação exata encontrada pela
       busca exaustiva (com poda), mesmo que os pagamentos estejam
       espalhados/longe uns dos outros.

Quando acha, pinta a(s) linha(s) do(s) pagamento(s) e a linha da OB com a
mesma cor, deixando visualmente claro quem está atrelado a quem.

Abas sem bloco de Pagamentos e/ou sem tabela OB/VALOR são ignoradas (nada é
alterado nelas).

Uso:
    python processar_pagamentos.py caminho_da_planilha.xlsx
"""

import argparse
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
]


# ---------------------------------------------------------------------------
# Helpers de localização (cada aba tem seu próprio layout, gerado pelo
# separar_tipo.py, então tudo é procurado dinamicamente em vez de assumir
# linha/coluna fixa)
# ---------------------------------------------------------------------------

def localizar_titulo(ws, prefixo):
    """Procura, na coluna A, a primeira linha cujo valor comece com
    `prefixo` (ex.: 'PAGAMENTOS', que aparece como 'PAGAMENTOS (47)').
    Devolve o número da linha ou None."""
    prefixo = prefixo.strip().upper()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).strip().upper().startswith(prefixo):
            return r
    return None


def localizar_bloco_ob(ws):
    """Procura em toda a aba uma célula 'OB' seguida, na coluna seguinte e
    mesma linha, por uma célula 'VALOR'. Devolve
    (col_ob, col_valor, linha_cabecalho) ou (None, None, None) se a aba não
    tiver essa tabela auxiliar."""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v and str(v).strip().upper() == "OB":
                v_valor = ws.cell(row=r, column=c + 1).value
                if v_valor and str(v_valor).strip().upper() == "VALOR":
                    return c, c + 1, r
    return None, None, None


def _ler_linhas_ob(ws, col_ob, col_valor, linha_inicial):
    """Lê (ob, valor) a partir de `linha_inicial`, parando na primeira linha
    totalmente vazia nas duas colunas, ou na primeira linha em que a coluna
    VALOR contenha algo que não seja número (sinal de que a tabela OB/VALOR
    terminou e começou outra tabela colada em seguida, sem linha vazia de
    separação - caso real observado na aba "07.08.25", onde uma tabela de
    "DESPESAS / RESPONSABILIDADE" vem logo abaixo)."""
    linhas = []
    r = linha_inicial
    while True:
        v_ob = ws.cell(row=r, column=col_ob).value
        v_valor = ws.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            break
        if v_valor is not None and not isinstance(v_valor, (int, float)):
            break
        linhas.append((v_ob, v_valor))
        r += 1
    return linhas


# ---------------------------------------------------------------------------
# Passo 1: remover OBs com VALOR = 0 (aba por aba)
# ---------------------------------------------------------------------------

def remover_valor_zero_aba(ws):
    """Remove as OBs com VALOR = 0 dentro dessa aba. Devolve
    (removidas, restantes) ou None se a aba não tiver bloco OB/VALOR."""
    col_ob, col_valor, linha_header = localizar_bloco_ob(ws)
    if col_ob is None:
        return None

    linha_inicial = linha_header + 1
    linhas = _ler_linhas_ob(ws, col_ob, col_valor, linha_inicial)

    total_antes = len(linhas)
    linhas_filtradas = [(ob, valor) for ob, valor in linhas if valor != 0]
    removidas = total_antes - len(linhas_filtradas)

    # limpa toda a área antiga do bloco (valor, borda e preenchimento)
    ultima_linha_bloco = linha_inicial + total_antes
    for r in range(linha_inicial, max(ws.max_row, ultima_linha_bloco) + 1):
        for col in (col_ob, col_valor):
            cell = ws.cell(row=r, column=col)
            cell.value = None
            cell.border = NO_BORDER
            cell.fill = NO_FILL

    # reescreve só as linhas que sobraram
    for i, (ob, valor) in enumerate(linhas_filtradas):
        r = linha_inicial + i
        c_ob = ws.cell(row=r, column=col_ob, value=ob)
        c_ob.font = NORMAL_FONT
        c_ob.border = BORDER
        c_ob.alignment = CENTER

        c_valor = ws.cell(row=r, column=col_valor, value=valor)
        c_valor.font = NORMAL_FONT
        c_valor.border = BORDER
        c_valor.number_format = "#,##0.00"
        c_valor.alignment = RIGHT

    return removidas, len(linhas_filtradas)


def _normalizar_abas(abas):
    """Normaliza a entrada de abas para uma lista de nomes."""
    if abas is None:
        return None
    if isinstance(abas, str):
        return [abas]
    return list(abas)


def remover_valor_zero(caminho, abas=None):
    wb = openpyxl.load_workbook(caminho)
    abas_selecionadas = _normalizar_abas(abas)
    abas_alvo = abas_selecionadas or wb.sheetnames

    total_removidas = 0
    abas_processadas = 0
    for nome in abas_alvo:
        if nome not in wb.sheetnames:
            print(f"Aba '{nome}' não encontrada. Ignorando.")
            continue

        ws = wb[nome]
        resultado = remover_valor_zero_aba(ws)
        if resultado is None:
            continue
        removidas, restantes = resultado
        abas_processadas += 1
        total_removidas += removidas
        print(f"[{nome}] OBs com VALOR = 0 removidas: {removidas} | restantes: {restantes}")

    wb.save(caminho)

    if abas_processadas == 0:
        print("Nenhuma aba com tabela OB/VALOR foi encontrada. Nada foi alterado.")
    else:
        print(f"\nTotal de OBs removidas: {total_removidas}, em {abas_processadas} aba(s).")


# ---------------------------------------------------------------------------
# Passo 2: conciliação pagamento(s) <-> OB (aba por aba)
# ---------------------------------------------------------------------------

def _centavos(valor):
    """Converte pra inteiro em centavos, pra comparar sem erro de ponto flutuante."""
    return round(float(valor) * 100)


def _linhas_pagamentos(ws, linha_titulo_pag):
    """Lê a tabela principal de Pagamentos (colunas A:G) dessa aba,
    começando logo após o título 'PAGAMENTOS (N)' e o cabeçalho de colunas
    que vem em seguida. Devolve [(linha, quantia)] enquanto a coluna
    'Linha' (A) tiver valor."""
    linhas = []
    r = linha_titulo_pag + 2  # +1 = cabeçalho de colunas, +2 = primeira linha de dado
    while ws.cell(row=r, column=1).value is not None:
        quantia = ws.cell(row=r, column=5).value  # coluna E = Quantia
        if quantia is not None:
            linhas.append((r, quantia))
        r += 1
    return linhas


def _linhas_ob_com_linha(ws, col_ob, col_valor, linha_inicial):
    """Como _ler_linhas_ob, mas guardando também o número da linha (usado
    pra pintar depois). Também para de ler se a coluna VALOR trouxer algo
    que não seja número (ver comentário em _ler_linhas_ob)."""
    linhas = []
    r = linha_inicial
    while True:
        v_ob = ws.cell(row=r, column=col_ob).value
        v_valor = ws.cell(row=r, column=col_valor).value
        if v_ob is None and v_valor is None:
            break
        if v_valor is not None and not isinstance(v_valor, (int, float)):
            break
        if v_valor is not None:
            linhas.append((r, v_ob, v_valor))
        r += 1
    return linhas


def _score_proximidade(combo):
    """Gera uma penalidade para combinações de 2+ pagamentos cujas linhas
    estão mais distantes entre si na planilha. Menor score é melhor. Dá
    peso extra para sequências contíguas (linhas vizinhas), que é o que
    queremos priorizar quando um único pagamento não bate exato com a OB.

    Usada apenas para desempate ENTRE combinações de 2+ itens - o caso de
    um único pagamento com valor exato é tratado à parte, antes disso, em
    `_buscar_subconjunto`."""
    linhas = sorted(linha for linha, _ in combo)
    span = linhas[-1] - linhas[0]
    gaps = [b - a for a, b in zip(linhas, linhas[1:])]
    contiguas = sum(1 for gap in gaps if gap == 1)
    return (span, sum(gaps), -contiguas, linhas[0])


def _buscar_subconjunto(pagamentos_disponiveis, alvo_centavos, max_tamanho):
    """Procura pagamento(s) cuja soma bate exatamente com `alvo_centavos`,
    seguindo esta ordem de prioridade:

        1) Match exato de UM único pagamento (retorna na hora, sem olhar
           para combinações de 2+ itens).
        2) Combinações de 2 ou mais pagamentos, priorizando as que usam
           pagamentos mais próximos entre si na planilha (idealmente linhas
           vizinhas / contíguas) - ver `_score_proximidade`.
        3) Se não houver combinação "próxima", qualquer combinação exata
           encontrada pela busca com poda serve (tentativa e erro geral,
           sem preferência de posição).

    Devolve a combinação (tupla de (linha, quantia)) ou None se nada bater.

    A busca em si usa poda (branch and bound) em vez de testar todas as
    combinações via itertools.combinations: os pagamentos são ordenados por
    valor, e a recursão corta um ramo assim que a soma parcial já passa do
    alvo (como a lista está em ordem crescente, os itens seguintes só
    aumentam a soma, então não há por que continuar). Sem isso, uma aba com
    ~50-80 pagamentos e várias OBs sem correspondência podia levar minutos
    (ou até travar na prática) tentando C(n, 6) combinações repetidamente."""
    itens = sorted(pagamentos_disponiveis, key=lambda x: _centavos(x[1]))
    valores = [_centavos(q) for _, q in itens]
    n = len(itens)
    max_tamanho = min(max_tamanho, n)

    # --- Prioridade 1: pagamento único com valor exatamente igual ao da OB ---
    for item, valor_centavos in zip(itens, valores):
        if valor_centavos == alvo_centavos:
            return (item,)

    # --- Prioridades 2 e 3: combinações de 2+ itens, do menor pro maior
    #     tamanho, com desempate por proximidade entre linhas ---
    melhor_combo = None
    melhor_score = None

    for tamanho in range(2, max_tamanho + 1):
        combo = _buscar_tamanho_fixo(itens, valores, alvo_centavos, tamanho)
        if combo is None:
            continue

        score = _score_proximidade(combo)
        if melhor_combo is None or score < melhor_score:
            melhor_combo = combo
            melhor_score = score

    return melhor_combo


def _buscar_tamanho_fixo(itens, valores, alvo, tamanho):
    """Busca com poda por uma combinação de exatamente `tamanho` itens (já
    ordenados por valor crescente em `itens`/`valores`) cuja soma bata com
    `alvo`. Devolve a tupla de itens ou None. Quando houver múltiplas
    combinações exatas desse mesmo tamanho, a melhor é escolhida pelo
    critério de proximidade entre linhas (`_score_proximidade`)."""
    n = len(itens)
    escolhidos = [0] * tamanho
    melhor_combo = None

    def rec(inicio, k, soma_parcial):
        nonlocal melhor_combo
        if k == tamanho:
            if soma_parcial == alvo:
                combo = tuple(itens[i] for i in escolhidos)
                if melhor_combo is None:
                    melhor_combo = combo
                else:
                    atual_score = _score_proximidade(combo)
                    melhor_score = _score_proximidade(melhor_combo)
                    if atual_score < melhor_score:
                        melhor_combo = combo
            return
        # não vale a pena tentar índices onde não sobram itens suficientes
        # para completar os `tamanho - k` restantes
        limite = n - (tamanho - k)
        for i in range(inicio, limite + 1):
            nova_soma = soma_parcial + valores[i]
            if nova_soma > alvo:
                # itens[i:] só tem valores >= valores[i] (lista ordenada),
                # então nenhum item a partir daqui pode mais servir
                break
            escolhidos[k] = i
            rec(i + 1, k + 1, nova_soma)
        return None

    rec(0, 0, 0)
    return melhor_combo


def conciliar_pagamentos_obs_aba(ws, max_combinacao=15):
    """Concilia pagamentos com OBs dentro de UMA aba. Devolve um dict com o
    relatório, ou None se a aba não tiver a estrutura esperada (bloco de
    Pagamentos e/ou tabela OB/VALOR)."""
    linha_titulo_pag = localizar_titulo(ws, "PAGAMENTOS")
    if linha_titulo_pag is None:
        return None

    col_ob, col_valor, linha_header_ob = localizar_bloco_ob(ws)
    if col_ob is None:
        return None

    pagamentos = _linhas_pagamentos(ws, linha_titulo_pag)
    obs = _linhas_ob_com_linha(ws, col_ob, col_valor, linha_header_ob + 1)

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
            for col in range(1, NUM_COLS_PRINCIPAL + 1):
                ws.cell(row=linha, column=col).fill = fill

    return {
        "aba": ws.title,
        "grupos": grupos,
        "obs_sem_match": obs_sem_match,
        "pagamentos_sem_ob": disponiveis,
        "total_obs": len(obs),
    }


def conciliar_pagamentos_obs(caminho, max_combinacao=6, abas=None):
    wb = openpyxl.load_workbook(caminho)
    abas_selecionadas = _normalizar_abas(abas)
    abas_alvo = abas_selecionadas or wb.sheetnames

    resultados = []
    for nome in abas_alvo:
        if nome not in wb.sheetnames:
            print(f"Aba '{nome}' não encontrada. Ignorando.")
            continue

        ws = wb[nome]
        resultado = conciliar_pagamentos_obs_aba(ws, max_combinacao=max_combinacao)
        if resultado is not None:
            resultados.append(resultado)

    wb.save(caminho)

    if not resultados:
        print("Nenhuma aba com bloco de Pagamentos + tabela OB/VALOR foi encontrada.")
        return

    for res in resultados:
        print(f"\n[{res['aba']}] OBs conciliadas: {len(res['grupos'])} de {res['total_obs']}")
        for grupo in res["grupos"]:
            n_pag = len(grupo["pagamento_rows"])
            composicao = "1 pagamento" if n_pag == 1 else f"{n_pag} pagamentos somados"
            print(f"  OB {grupo['ob_codigo']} (R$ {grupo['valor']:.2f}) <- {composicao}, linha(s) {grupo['pagamento_rows']}")

        if res["obs_sem_match"]:
            print(f"  ⚠️  ATENÇÃO: {len(res['obs_sem_match'])} OB(s) SEM correspondência encontrada nesta aba!")
            for ob_row, ob_codigo, valor in res["obs_sem_match"]:
                print(f"    OB {ob_codigo} (R$ {valor:.2f})")

        if res["pagamentos_sem_ob"]:
            print(f"  Pagamentos sem OB correspondente: {len(res['pagamentos_sem_ob'])}")
            for linha, quantia in res["pagamentos_sem_ob"]:
                print(f"    Linha {linha} (R$ {quantia:.2f})")

    # Resumo global: como toda OB deveria, em tese, ter um pagamento (ou
    # combinação de pagamentos) correspondente, qualquer OB sem match é um
    # sinal de que algo precisa ser conferido manualmente. Por isso esse
    # aviso fica bem destacado no final, reunindo todas as abas, e não só
    # espalhado aba por aba.
    total_obs_sem_match = sum(len(res["obs_sem_match"]) for res in resultados)
    if total_obs_sem_match > 0:
        print("\n" + "=" * 60)
        print(f"⚠️  ATENÇÃO: {total_obs_sem_match} OB(s) NÃO conciliada(s) no total!")
        print("=" * 60)
        for res in resultados:
            if res["obs_sem_match"]:
                for ob_row, ob_codigo, valor in res["obs_sem_match"]:
                    print(f"  [{res['aba']}] OB {ob_codigo} (R$ {valor:.2f}) - linha {ob_row}")
        print("=" * 60)
    else:
        print("\n✅ Todas as OBs foram conciliadas com sucesso em todas as abas.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Processa pagamentos e OBs em uma planilha Excel.")
    parser.add_argument("caminho", nargs="?", default="conciliacao.xlsx", help="Caminho da planilha Excel")
    parser.add_argument("--aba", "--abas", dest="abas", action="append", help="Nome da aba a ser processada. Pode ser informado mais de uma vez.")
    args = parser.parse_args()

    remover_valor_zero(args.caminho, abas=args.abas)
    print()
    conciliar_pagamentos_obs(args.caminho, abas=args.abas)