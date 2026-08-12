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
A busca por match segue uma ordem de prioridade:

    1) Match exato de UM único pagamento cujo valor bate com o de UMA OB.
       Feito primeiro, pra TODAS as OBs, sem nenhuma ambiguidade (não
       depende de olhar combinações).

    2) Entre as OBs que sobraram sem match no passo 1, tenta somar 2 ou
       mais delas entre si (ex.: duas remessas que na prática foram pagas
       juntas em um único lote) e ver se essa soma bate com um pagamento ou
       combinação de pagamentos ainda disponível - ver
       `_buscar_combo_obs_pagamentos`. Importante isso vir ANTES do passo
       3: se uma OB isolada "coubesse" dentro de um pool de pagamentos de
       valor repetido (ex.: vários pagamentos de R$300,00 iguais), o passo
       3 casaria essa OB sozinha com um subconjunto arbitrário desse pool,
       "gastando" pagamentos que na verdade deveriam ficar reservados para
       serem somados junto com OUTRA OB (ex.: duas OBs que juntas fecham
       o valor de TODO o pool de repetidos). Rodar a combinação de OBs
       primeiro evita esse tipo de match "ganancioso" e incompleto.

    3) Para as OBs que ainda sobrarem, combinações de 2 ou mais pagamentos
       cuja soma bate com o valor de UMA OB isolada, dando preferência para
       pagamentos que estejam fisicamente próximos entre si na planilha
       (idealmente na linha logo abaixo ou logo acima um do outro).

    4) Se nada acima servir, vale qualquer combinação exata encontrada pela
       busca exaustiva (com poda), mesmo que os pagamentos estejam
       espalhados/longe uns dos outros.

Quando acha, pinta a(s) linha(s) do(s) pagamento(s) e a(s) linha(s) da(s)
OB(s) com a mesma cor, deixando visualmente claro quem está atrelado a
quem (inclusive quando o grupo tem mais de uma OB).

Abas sem bloco de Pagamentos e/ou sem tabela OB/VALOR são ignoradas (nada é
alterado nelas).

Passo 3 (opcional, roda só se --extratos for informado): para os pagamentos
que sobraram sem OB correspondente no Passo 2, cruza cada um com a tabela de
PAGAMENTOS da planilha de extratos (saida.xlsx, gerada por outro script a
partir dos extratos bancários), usando como chave a combinação Data
Transação + Quantia (comparado com Data + Valor (R$) do extrato). Para cada
pagamento sem OB:

    1) Se achar uma linha do extrato com a mesma Data+Valor e o Histórico
       dela começar com "Tarifa", preenche a coluna "Status" com "TARIFA".
    2) Se achar e não for tarifa, preenche "Status" com o texto do
       Histórico do extrato e a coluna "Justificativas - Não reconciliadas"
       com "MANDAR PARA CONTAS A PAGAR".
    3) Se não achar nenhuma linha do extrato com aquela Data+Valor, a linha
       é deixada como está (fica sinalizada no log para conferência
       manual).

Se mais de uma linha do extrato bater com a mesma Data+Valor (ambíguo), a
primeira encontrada é usada, mas o caso é destacado no log para revisão.

Uso:
    python processar_pagamentos.py caminho_da_planilha.xlsx
    python processar_pagamentos.py caminho_da_planilha.xlsx --extratos saida.xlsx
"""

import argparse
import datetime
import itertools
import math
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


# trava de segurança: para um dado tamanho de combinação, se o número de
# combinações possíveis (n escolhe k) passar disso, esse tamanho é pulado em
# vez de arriscar travar o script numa aba com muitos itens sem match
LIMITE_COMBINACOES = 200_000


def _enumerar_somas(itens, valor_fn, max_tamanho, teto_centavos=None, limite=LIMITE_COMBINACOES):
    """Enumera, com poda, combinações de `itens` de tamanho 1 até
    `max_tamanho`, devolvendo um mapa {soma_centavos: combinação}. `valor_fn`
    extrai o valor monetário de um item. Os itens são ordenados por valor
    crescente e percorridos em ordem crescente de TAMANHO de combinação, então
    pra cada soma a combinação guardada é sempre a menor/mais enxuta
    encontrada primeiro (`mapa.setdefault` não sobrescreve).

    Se `teto_centavos` for informado, poda qualquer ramo cuja soma parcial já
    ultrapasse esse teto - usado aqui pra não perder tempo somando
    combinações de pagamentos maiores que qualquer soma de OBs que faça
    sentido consultar depois. `limite` é a mesma trava de `LIMITE_COMBINACOES`:
    se o número de combinações possíveis (n escolhe k) for grande demais pra
    um dado tamanho, esse tamanho é pulado."""
    ordenado = sorted(itens, key=lambda it: _centavos(valor_fn(it)))
    valores = [_centavos(valor_fn(it)) for it in ordenado]
    n = len(ordenado)
    mapa = {}

    for tamanho in range(1, min(max_tamanho, n) + 1):
        if math.comb(n, tamanho) > limite:
            continue
        escolhidos = [0] * tamanho

        def rec(inicio, k, soma_parcial):
            if k == tamanho:
                mapa.setdefault(soma_parcial, tuple(ordenado[i] for i in escolhidos))
                return
            limite_i = n - (tamanho - k)
            for i in range(inicio, limite_i + 1):
                nova_soma = soma_parcial + valores[i]
                if teto_centavos is not None and nova_soma > teto_centavos:
                    # itens[i:] só tem valores >= valores[i] (lista ordenada),
                    # então nenhum item a partir daqui cabe mais no teto
                    break
                escolhidos[k] = i
                rec(i + 1, k + 1, nova_soma)

        rec(0, 0, 0)

    return mapa


def _buscar_combo_obs_pagamentos(obs_disponiveis, pagamentos_disponiveis, max_tamanho_ob, max_tamanho_pag):
    """Última tentativa de conciliação (prioridade 4, ver docstring do
    módulo): quando uma OB isolada não bate com nenhum pagamento nem
    combinação de pagamentos, tenta somar 2 ou mais OBs entre si (dentre as
    que já sobraram sem match) e ver se essa soma bate com um pagamento ou
    combinação de pagamentos ainda disponível - caso real de várias remessas
    (OBs) que na prática saíram juntas num único lote/pagamento.

    Para não repetir uma busca combinatória cara a cada combinação de OBs
    testada, as somas possíveis do lado dos pagamentos são pré-computadas
    UMA VEZ (`_enumerar_somas`), e cada combinação de OBs só faz uma consulta
    O(1) nesse mapa em vez de refazer a busca inteira do zero. Isso troca
    "combinações de OBs" x "busca exaustiva de pagamentos" (multiplicativo,
    caro) por "combinações de OBs" + "combinações de pagamentos" (aditivo).

    Diferente da prioridade 1-3 (que prioriza pagamentos fisicamente
    próximos entre si via `_score_proximidade`), esta etapa não aplica esse
    desempate por proximidade - prioriza apenas o menor número de itens dos
    dois lados, o que já cobre o caso típico (poucas OBs somadas contra
    poucos pagamentos somados).

    `obs_disponiveis` é uma lista de (linha, codigo, valor); a busca
    considera combinações de tamanho `2` até `max_tamanho_ob` (limitado ao
    número de OBs disponíveis). Devolve (combo_obs, combo_pagamentos) ou
    None se nada bater. combo_obs é uma tupla de (linha, codigo, valor);
    combo_pagamentos é uma tupla de (linha, quantia)."""
    obs_ordenadas = sorted(obs_disponiveis, key=lambda x: x[2], reverse=True)
    n_obs = len(obs_ordenadas)
    max_tamanho_ob = min(max_tamanho_ob, n_obs)
    if max_tamanho_ob < 2 or not pagamentos_disponiveis:
        return None

    # teto: nenhuma combinação de OBs testada aqui passa da soma das
    # `max_tamanho_ob` maiores OBs disponíveis - usado pra podar a
    # enumeração de somas de pagamentos e não desperdiçar tempo somando
    # combinações de pagamentos maiores que qualquer alvo plausível
    teto_centavos = sum(_centavos(v) for _, _, v in obs_ordenadas[:max_tamanho_ob])

    mapa_pagamentos = _enumerar_somas(
        pagamentos_disponiveis, valor_fn=lambda p: p[1],
        max_tamanho=max_tamanho_pag, teto_centavos=teto_centavos,
    )
    if not mapa_pagamentos:
        return None

    for tamanho in range(2, max_tamanho_ob + 1):
        if math.comb(n_obs, tamanho) > LIMITE_COMBINACOES:
            # combinatória grande demais pra esse tamanho de combo de OBs -
            # pula pro próximo tamanho em vez de arriscar travar
            continue
        for combo_obs in itertools.combinations(obs_ordenadas, tamanho):
            alvo = sum(_centavos(valor) for _, _, valor in combo_obs)
            combo_pagamentos = mapa_pagamentos.get(alvo)
            if combo_pagamentos is not None:
                return combo_obs, combo_pagamentos

    return None


def conciliar_pagamentos_obs_aba(ws, max_combinacao=15, max_combinacao_obs=6):
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

    # --- Prioridade 1: match exato de UM único pagamento, pra TODAS as OBs
    #     primeiro (sem ambiguidade, não depende de nenhuma combinação) ---
    obs_pendentes = []
    for ob_row, ob_codigo, valor in obs_ordenadas:
        alvo = _centavos(valor)
        pagamento_exato = next(
            (p for p in disponiveis if _centavos(p[1]) == alvo), None
        )
        if pagamento_exato is None:
            obs_pendentes.append((ob_row, ob_codigo, valor))
            continue
        grupos.append({
            "ob_rows": [ob_row],
            "ob_codigos": [ob_codigo],
            "valor": valor,
            "pagamento_rows": [pagamento_exato[0]],
        })
        disponiveis.remove(pagamento_exato)

    # --- Prioridade 2: combinações de 2+ OBs (dentre as que sobraram) vs
    #     combinações de pagamentos - ver comentário na docstring do módulo
    #     sobre por que isso roda ANTES da prioridade 3. Repete a busca até
    #     não achar mais nenhuma combinação, já que cada match muda o que
    #     ainda está disponível dos dois lados. ---
    while len(obs_pendentes) >= 2 and disponiveis:
        resultado_combo = _buscar_combo_obs_pagamentos(
            obs_pendentes, disponiveis, max_combinacao_obs, max_combinacao
        )
        if resultado_combo is None:
            break

        combo_obs, combo_pagamentos = resultado_combo
        ob_rows = [ob_row for ob_row, _, _ in combo_obs]
        ob_codigos = [ob_codigo for _, ob_codigo, _ in combo_obs]
        valor_total = sum(valor for _, _, valor in combo_obs)
        linhas_pagamento = [linha for linha, _ in combo_pagamentos]

        grupos.append({
            "ob_rows": ob_rows,
            "ob_codigos": ob_codigos,
            "valor": valor_total,
            "pagamento_rows": linhas_pagamento,
        })

        usados_obs = set(ob_rows)
        obs_pendentes = [o for o in obs_pendentes if o[0] not in usados_obs]
        usados_pag = set(linhas_pagamento)
        disponiveis = [p for p in disponiveis if p[0] not in usados_pag]

    # --- Prioridades 3 e 4: pra quem ainda sobrou, OB isolada vs combinação
    #     de 2+ pagamentos (`_buscar_subconjunto` já cuida internamente de
    #     priorizar pagamentos próximos entre si antes de aceitar qualquer
    #     combinação exaustiva) ---
    obs_sem_match = []
    for ob_row, ob_codigo, valor in obs_pendentes:
        alvo = _centavos(valor)
        combo = _buscar_subconjunto(disponiveis, alvo, max_combinacao)
        if combo is None:
            obs_sem_match.append((ob_row, ob_codigo, valor))
            continue
        linhas_pagamento = [linha for linha, _ in combo]
        grupos.append({
            "ob_rows": [ob_row],
            "ob_codigos": [ob_codigo],
            "valor": valor,
            "pagamento_rows": linhas_pagamento,
        })
        usados = set(linhas_pagamento)
        disponiveis = [p for p in disponiveis if p[0] not in usados]

    # pinta os grupos conciliados
    for i, grupo in enumerate(grupos):
        cor = CORES_CONCILIACAO[i % len(CORES_CONCILIACAO)]
        fill = PatternFill("solid", fgColor=cor)

        for ob_row in grupo["ob_rows"]:
            for col in (col_ob, col_valor):
                ws.cell(row=ob_row, column=col).fill = fill

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


def conciliar_pagamentos_obs(caminho, max_combinacao=6, max_combinacao_obs=6, abas=None):
    wb = openpyxl.load_workbook(caminho)
    abas_selecionadas = _normalizar_abas(abas)
    abas_alvo = abas_selecionadas or wb.sheetnames

    resultados = []
    for nome in abas_alvo:
        if nome not in wb.sheetnames:
            print(f"Aba '{nome}' não encontrada. Ignorando.")
            continue

        ws = wb[nome]
        resultado = conciliar_pagamentos_obs_aba(
            ws, max_combinacao=max_combinacao, max_combinacao_obs=max_combinacao_obs
        )
        if resultado is not None:
            resultados.append(resultado)

    wb.save(caminho)

    if not resultados:
        print("Nenhuma aba com bloco de Pagamentos + tabela OB/VALOR foi encontrada.")
        return

    for res in resultados:
        n_obs_conciliadas = sum(len(grupo["ob_rows"]) for grupo in res["grupos"])
        print(f"\n[{res['aba']}] OBs conciliadas: {n_obs_conciliadas} de {res['total_obs']} (em {len(res['grupos'])} grupo(s))")
        for grupo in res["grupos"]:
            n_pag = len(grupo["pagamento_rows"])
            composicao = "1 pagamento" if n_pag == 1 else f"{n_pag} pagamentos somados"
            obs_str = " + ".join(str(c) for c in grupo["ob_codigos"])
            print(f"  OB(s) {obs_str} (R$ {grupo['valor']:.2f} total) <- {composicao}, linha(s) {grupo['pagamento_rows']}")

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


# ---------------------------------------------------------------------------
# Passo 3: para pagamentos que sobraram SEM conciliação (nenhuma OB bateu),
# cruzar com a tabela de extratos (saida.xlsx) por Data + Valor e preencher
# Status / Justificativas.
# ---------------------------------------------------------------------------

MESES_PT = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}

COL_STATUS = 6
COL_JUSTIFICATIVA = 7
JUSTIFICATIVA_CONTAS_A_PAGAR = "MANDAR PARA CONTAS A PAGAR"


def _data_para_tupla(valor):
    """Aceita 'dd/mmm/aa' (ex 27/jul/26), 'dd/mm/aaaa' (ex 27/07/2026) e
    também células de data reais do Excel (datetime.datetime / datetime.date,
    como vem em algumas abas de conciliacao.xlsx). Devolve (ano, mes, dia),
    ou None se não conseguir interpretar."""
    if valor is None:
        return None

    if isinstance(valor, datetime.datetime):
        return (valor.year, valor.month, valor.day)
    if isinstance(valor, datetime.date):
        return (valor.year, valor.month, valor.day)

    partes = str(valor).strip().split("/")
    if len(partes) != 3:
        return None
    dia_str, mes_str, ano_str = partes
    try:
        dia = int(dia_str)
    except ValueError:
        return None

    if mes_str.isdigit():
        mes = int(mes_str)
    else:
        mes = MESES_PT.get(mes_str.strip().lower()[:3])
        if mes is None:
            return None

    try:
        ano = int(ano_str)
    except ValueError:
        return None
    if ano < 100:
        ano += 2000

    return (ano, mes, dia)


def _ler_pagamentos_extrato_aba(ws):
    """Lê a tabela de PAGAMENTOS de uma aba de saida.xlsx (colunas
    Data/Histórico/Documento/Valor). Devolve [(data, histórico, valor)]."""
    linha_titulo = localizar_titulo(ws, "PAGAMENTOS")
    if linha_titulo is None:
        return []

    linhas = []
    r = linha_titulo + 2  # +1 cabeçalho de colunas, +2 primeira linha de dado
    while True:
        data = ws.cell(row=r, column=1).value
        if data is None:
            break
        historico = ws.cell(row=r, column=2).value
        valor = ws.cell(row=r, column=4).value
        if isinstance(valor, (int, float)):
            linhas.append((data, historico, valor))
        r += 1
    return linhas


def _indexar_extrato(caminho_extratos):
    """Monta um índice (ano, mes, dia, valor_em_centavos) -> [(aba,
    histórico), ...] a partir de todas as abas de PAGAMENTOS em
    saida.xlsx."""
    wb = openpyxl.load_workbook(caminho_extratos, data_only=True)
    indice = {}
    for nome in wb.sheetnames:
        ws = wb[nome]
        for data, historico, valor in _ler_pagamentos_extrato_aba(ws):
            chave_data = _data_para_tupla(data)
            if chave_data is None:
                continue
            chave = (*chave_data, _centavos(valor))
            indice.setdefault(chave, []).append((nome, historico))
    return indice


def atualizar_status_sem_conciliacao_aba(ws, pagamentos_sem_ob, indice_extrato):
    """Para cada pagamento sem OB correspondente (linha, quantia), procura a
    mesma Data+Valor na tabela de extratos:
      - se achar e o Histórico começar com 'Tarifa' -> Status = 'TARIFA'
      - se achar e não for tarifa -> Status = Histórico do extrato,
        Justificativas = 'MANDAR PARA CONTAS A PAGAR'
      - se não achar nenhuma correspondência -> nada é alterado (fica para
        conferência manual).
    Devolve um relatório (listas de linhas tratadas como tarifa, mandadas
    pra contas a pagar, sem correspondência no extrato e ambíguas)."""
    tratados_tarifa = []
    tratados_contas_a_pagar = []
    sem_correspondencia = []
    ambiguos = []

    for linha, quantia in pagamentos_sem_ob:
        data_transacao = ws.cell(row=linha, column=4).value
        chave_data = _data_para_tupla(data_transacao)
        if chave_data is None:
            sem_correspondencia.append((linha, quantia))
            continue

        chave = (*chave_data, _centavos(quantia))
        candidatos = indice_extrato.get(chave)
        if not candidatos:
            sem_correspondencia.append((linha, quantia))
            continue

        if len(candidatos) > 1:
            ambiguos.append((linha, quantia, candidatos))

        aba_origem, historico = candidatos[0]
        historico_str = str(historico or "").strip()

        if historico_str.lower().startswith("tarifa"):
            ws.cell(row=linha, column=COL_STATUS, value="TARIFA")
            tratados_tarifa.append((linha, quantia, aba_origem, historico_str))
        else:
            ws.cell(row=linha, column=COL_STATUS, value=historico_str)
            ws.cell(row=linha, column=COL_JUSTIFICATIVA, value=JUSTIFICATIVA_CONTAS_A_PAGAR)
            tratados_contas_a_pagar.append((linha, quantia, aba_origem, historico_str))

    return {
        "tarifa": tratados_tarifa,
        "contas_a_pagar": tratados_contas_a_pagar,
        "sem_correspondencia": sem_correspondencia,
        "ambiguos": ambiguos,
    }


def atualizar_status_sem_conciliacao(caminho_conciliacao, caminho_extratos, max_combinacao=6, max_combinacao_obs=6, abas=None):
    """Roda a conciliação normal (pagamento <-> OB) e, para o que sobrar sem
    match, cruza com saida.xlsx por Data+Valor para preencher Status /
    Justificativas. Salva o resultado em `caminho_conciliacao`."""
    indice_extrato = _indexar_extrato(caminho_extratos)

    wb = openpyxl.load_workbook(caminho_conciliacao)
    abas_selecionadas = _normalizar_abas(abas)
    abas_alvo = abas_selecionadas or wb.sheetnames

    for nome in abas_alvo:
        if nome not in wb.sheetnames:
            print(f"Aba '{nome}' não encontrada em {caminho_conciliacao}. Ignorando.")
            continue

        ws = wb[nome]
        resultado_conciliacao = conciliar_pagamentos_obs_aba(
            ws, max_combinacao=max_combinacao, max_combinacao_obs=max_combinacao_obs
        )
        if resultado_conciliacao is None:
            continue

        n_obs_conciliadas = sum(len(grupo["ob_rows"]) for grupo in resultado_conciliacao["grupos"])
        print(f"\n[{nome}] OBs conciliadas: {n_obs_conciliadas} de {resultado_conciliacao['total_obs']} (em {len(resultado_conciliacao['grupos'])} grupo(s))")

        relatorio = atualizar_status_sem_conciliacao_aba(
            ws, resultado_conciliacao["pagamentos_sem_ob"], indice_extrato
        )

        print(f"\n[{nome}] Pagamentos sem conciliação (OB): {len(resultado_conciliacao['pagamentos_sem_ob'])}")
        print(f"  -> Marcados como TARIFA: {len(relatorio['tarifa'])}")
        for linha, quantia, aba_origem, historico in relatorio["tarifa"]:
            print(f"     linha {linha} (R$ {quantia:.2f}) <- [{aba_origem}] {historico}")

        print(f"  -> Marcados p/ CONTAS A PAGAR: {len(relatorio['contas_a_pagar'])}")
        for linha, quantia, aba_origem, historico in relatorio["contas_a_pagar"]:
            print(f"     linha {linha} (R$ {quantia:.2f}) <- [{aba_origem}] {historico}")

        if relatorio["sem_correspondencia"]:
            print(f"  ⚠️  Sem correspondência nenhuma no extrato: {len(relatorio['sem_correspondencia'])}")
            for linha, quantia in relatorio["sem_correspondencia"]:
                print(f"     linha {linha} (R$ {quantia:.2f}) - conferir manualmente")

        if relatorio["ambiguos"]:
            print(f"  ⚠️  Data+Valor batendo com MAIS de uma linha do extrato (usada a primeira encontrada): {len(relatorio['ambiguos'])}")
            for linha, quantia, candidatos in relatorio["ambiguos"]:
                print(f"     linha {linha} (R$ {quantia:.2f}) candidatos: {candidatos}")

    wb.save(caminho_conciliacao)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Processa pagamentos e OBs em uma planilha Excel.")
    parser.add_argument("caminho", nargs="?", default="conciliacao.xlsx", help="Caminho da planilha Excel")
    parser.add_argument("--aba", "--abas", dest="abas", action="append", help="Nome da aba a ser processada. Pode ser informado mais de uma vez.")
    parser.add_argument("--extratos", dest="extratos", default=None, help="Caminho da planilha de extratos (saida.xlsx). Se informado, roda também o passo 3 (marcar TARIFA / CONTAS A PAGAR nos pagamentos sem OB).")
    parser.add_argument("--max-combinacao-obs", dest="max_combinacao_obs", type=int, default=6, help="Tamanho máximo de combinação de OBs (2 ou mais) testada na prioridade 4, quando OBs isoladas não batem com nenhum pagamento. Padrão: 6.")
    args = parser.parse_args()

    remover_valor_zero(args.caminho, abas=args.abas)
    print()

    if args.extratos:
        # já conciliamos por OB dentro de atualizar_status_sem_conciliacao,
        # então não precisa chamar conciliar_pagamentos_obs de novo aqui.
        atualizar_status_sem_conciliacao(
            args.caminho, args.extratos, max_combinacao_obs=args.max_combinacao_obs, abas=args.abas
        )
    else:
        conciliar_pagamentos_obs(args.caminho, max_combinacao_obs=args.max_combinacao_obs, abas=args.abas)