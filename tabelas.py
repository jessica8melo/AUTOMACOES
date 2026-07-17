#!/usr/bin/env python3
"""
Busca campos específicos dentro de uma planilha xlsx (ex.: Formulário de
Qualidade - FQ), localizando automaticamente a linha de cabeçalho da tabela
e lendo os valores da(s) linha(s) de dados logo abaixo.

Uso:
    python buscar_campos_xlsx.py caminho/do/arquivo.xlsx
"""

import sys
import re
import unicodedata
from difflib import SequenceMatcher

import openpyxl

# ---------------------------------------------------------------------------
# Campos que devem ser buscados na planilha.
# Adicione, remova ou edite livremente.
# ---------------------------------------------------------------------------
CAMPOS_PROCURADOS = [
    "Unidade de Medida",
    "Quantidade do Item",
    "Valor Preço Unitário",
    "UOR",
    "Conta Contábil",
]


# ---------------------------------------------------------------------------
# Utilidades de normalização e comparação "aproximada" de texto
# (mesma lógica usada no script de busca em PDF)
# ---------------------------------------------------------------------------
def normalizar(texto: str) -> str:
    if not texto:
        return ""
    texto = re.sub(r"\([^)]*\)", " ", texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def parecido(a: str, b: str, limite: float = 0.72) -> bool:
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return False
    if na == nb or na in nb or nb in na:
        return True
    return SequenceMatcher(None, na, nb).ratio() >= limite


def valor_vazio(valor) -> bool:
    """True se a célula estiver vazia ou contiver só espaços em branco."""
    return valor is None or (isinstance(valor, str) and not valor.strip())


# ---------------------------------------------------------------------------
# Localização da tabela dentro da planilha
# ---------------------------------------------------------------------------
def encontrar_linha_cabecalho(ws):
    """
    Percorre as linhas da planilha e devolve a que mais parece um cabeçalho
    de tabela: linha com várias células de texto que batem (aproximadamente)
    com os nomes dos campos procurados.
    """
    melhor_linha = None
    melhor_pontuacao = 0

    for linha in ws.iter_rows(min_row=1, max_row=ws.max_row):
        celulas_texto = [c for c in linha if isinstance(c.value, str) and c.value.strip()]
        if len(celulas_texto) < 3:
            continue
        pontuacao = sum(
            1
            for celula in celulas_texto
            for campo in CAMPOS_PROCURADOS
            if parecido(celula.value, campo)
        )
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_linha = linha

    return melhor_linha, melhor_pontuacao


def mapear_colunas(linha_cabecalho):
    """Devolve {indice_da_coluna: texto_do_cabecalho} para a linha dada."""
    return {
        celula.column: celula.value.strip()
        for celula in linha_cabecalho
        if isinstance(celula.value, str) and celula.value.strip()
    }


def extrair_linhas_de_dados(ws, linha_cabecalho, colunas):
    """
    A partir da linha seguinte ao cabeçalho, coleta as linhas de dados da
    tabela. Pula linhas em branco/whitespace (comuns como separador visual)
    e para ao encontrar uma linha de total (ex.: 'Total') ou o fim da tabela.
    """
    linha_inicio = linha_cabecalho[0].row + 1
    linhas_dados = []

    for num_linha in range(linha_inicio, ws.max_row + 1):
        valores = {col: ws.cell(row=num_linha, column=col).value for col in colunas}

        primeira_col = min(colunas.keys())
        primeira_celula = valores.get(primeira_col)

        # Linha de total encerra a tabela
        if isinstance(primeira_celula, str) and parecido(primeira_celula, "Total"):
            break

        # Linha totalmente em branco: pula (pode ser um separador visual)
        if all(valor_vazio(v) for v in valores.values()):
            continue

        linhas_dados.append(valores)

    return linhas_dados


def buscar_campo(campo, colunas, linhas_dados):
    """Encontra a coluna cujo cabeçalho é parecido com `campo` e devolve os
    valores dessa coluna em cada linha de dados."""
    coluna_encontrada = None
    for col_idx, texto_cabecalho in colunas.items():
        if parecido(texto_cabecalho, campo):
            coluna_encontrada = col_idx
            break

    if coluna_encontrada is None:
        return None

    valores = []
    for linha in linhas_dados:
        v = linha.get(coluna_encontrada)
        if isinstance(v, str):
            v = v.strip()
        valores.append(v)

    return valores


# ---------------------------------------------------------------------------
# Função reutilizável (chamada pelo main.py e também usável via CLI)
# ---------------------------------------------------------------------------
def processar_xlsx(caminho_xlsx: str) -> dict:
    """
    Analisa a planilha em `caminho_xlsx` e devolve um dicionário
    {campo: lista_de_valores_ou_None}, um valor por linha de dados da tabela
    encontrada na primeira aba onde ela aparecer.
    Lança exceção se o arquivo não puder ser aberto/lido.
    """
    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)

    for ws in wb.worksheets:
        linha_cabecalho, pontuacao = encontrar_linha_cabecalho(ws)
        if not linha_cabecalho or pontuacao == 0:
            continue  # esta aba não parece ter a tabela procurada

        colunas = mapear_colunas(linha_cabecalho)
        linhas_dados = extrair_linhas_de_dados(ws, linha_cabecalho, colunas)

        resultado = {
            "_aba": ws.title,
            "_linha_cabecalho": linha_cabecalho[0].row,
            "_num_linhas_dados": len(linhas_dados),
        }
        for campo in CAMPOS_PROCURADOS:
            resultado[campo] = buscar_campo(campo, colunas, linhas_dados)

        return resultado

    return None  # nenhuma tabela reconhecível em nenhuma aba


def imprimir_resultado(caminho_xlsx: str, resultado: dict) -> None:
    print(f"Planilha analisada: {caminho_xlsx}\n")

    if resultado is None:
        print("[ERRO] Não foi encontrada nenhuma tabela reconhecível na planilha.")
        return

    print(f"[Aba: {resultado['_aba']}] tabela encontrada na linha {resultado['_linha_cabecalho']}, "
          f"{resultado['_num_linhas_dados']} linha(s) de dados.\n")

    for campo in CAMPOS_PROCURADOS:
        valores = resultado.get(campo)
        if valores is None:
            print(f"[ERRO] Não foi possível encontrar o campo '{campo}' na planilha.")
        elif len(valores) == 1:
            print(f"[SUCESSO] Campo '{campo}' encontrado. Valor: {valores[0]}")
        else:
            print(f"[SUCESSO] Campo '{campo}' encontrado ({len(valores)} linhas). "
                  f"Valores: {valores}")


# ---------------------------------------------------------------------------
# Programa principal (uso via linha de comando, standalone)
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python tabelas.py caminho/do/arquivo.xlsx")
        sys.exit(1)

    caminho_xlsx = sys.argv[1]

    try:
        resultado = processar_xlsx(caminho_xlsx)
    except Exception as erro:
        print(f"[ERRO] Não foi possível abrir o arquivo '{caminho_xlsx}': {erro}")
        sys.exit(1)

    imprimir_resultado(caminho_xlsx, resultado)


if __name__ == "__main__":
    main()