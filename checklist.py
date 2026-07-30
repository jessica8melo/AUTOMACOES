#!/usr/bin/env python3
"""
Ponto de entrada do fluxo "checklist": recebe o documento de checklist que
está sendo aplicado (ex.: o .docx "FQ415-031_v10 - Checklist OC Padrão –
Com Contrato.docx"), identifica a QUAL fluxo ele pertence (fluxos.FLUXOS) e,
para cada documento atrelado a esse checklist (fluxos.documentos_do_fluxo),
localiza automaticamente o(s) arquivo(s) correspondente(s) dentro da pasta
de anexos (por padrão "Anexos/") e chama pdfs.py/tabelas.py para extrair
os campos daquele documento (fluxos.campos_do_documento -> documentos.py)
— em vez de só listar quais campos seriam esperados.

A identificação do fluxo é feita pelo código (ex.: "FQ415-031") presente
no NOME do arquivo de checklist; se não achar no nome, tenta no TEXTO do
.docx (lido direto do zip, sem depender de bibliotecas externas).

A identificação de qual anexo corresponde a qual documento do fluxo reusa
a mesma lógica de main.py (detectar_tipo/extrair_texto_para_identificacao)
e doc_types.identificar_documento, restringindo os candidatos aos
documentos daquele fluxo específico.

Uso:
    python checklist.py caminho/da/pasta                        # forma recomendada:
                                                                  # o script acha sozinho o
                                                                  # arquivo de checklist dentro
                                                                  # da pasta e usa essa MESMA
                                                                  # pasta como anexos
    python checklist.py "Checklists/FQ415-031_v10 - Checklist OC Padrão – Com Contrato.docx"
    python checklist.py FQ415-031          # também aceita o código direto
    python checklist.py FQ415-031 caminho/da/pasta/de/anexos   # pasta alternativa (padrão: "Anexos")
"""

import os
import re
import sys
import zipfile

import doc_types
import fluxos
import main as automacao_main
import pdfs
import tabelas

PADRAO_CODIGO_FLUXO = re.compile(r"FQ\s*\d{3}[\s_-]*\d{3}", re.IGNORECASE)

PASTA_ANEXOS_PADRAO = "Anexos"


def _normalizar_codigo(bruto: str) -> str:
    """'fq 415_031', 'FQ415 031' etc. -> 'FQ415-031'."""
    bruto = bruto.upper()
    numeros = re.findall(r"\d+", bruto)
    if len(numeros) < 2:
        return bruto.strip()
    return f"FQ{numeros[0]}-{numeros[1]}"


def extrair_texto_docx(caminho: str) -> str:
    """Extrai o texto bruto de um .docx lendo word/document.xml direto do
    zip (um .docx é um arquivo zip), sem depender do python-docx."""
    try:
        with zipfile.ZipFile(caminho) as z:
            with z.open("word/document.xml") as f:
                xml = f.read().decode("utf-8", errors="ignore")
    except Exception:
        return ""
    return re.sub(r"<[^>]+>", " ", xml)


def identificar_fluxo(caminho: str) -> str:
    """Identifica o código do fluxo (ex.: 'FQ415-031') a partir do NOME do
    arquivo de checklist e, se não achar, tenta pelo TEXTO do documento.
    Devolve o código normalizado se ele existir em fluxos.FLUXOS, ou None."""
    nome = os.path.basename(caminho)
    match = PADRAO_CODIGO_FLUXO.search(nome)

    if not match and caminho.lower().endswith(".docx"):
        match = PADRAO_CODIGO_FLUXO.search(extrair_texto_docx(caminho))

    if not match:
        return None

    codigo = _normalizar_codigo(match.group(0))
    return codigo if codigo in fluxos.FLUXOS else None


def _listar_arquivos_pasta(pasta: str, excluir: set = None) -> list:
    """Lista (em ordem alfabética) os arquivos dentro de `pasta`, ignorando
    subpastas, arquivos ocultos e os caminhos (absolutos) presentes em
    `excluir` (por exemplo, o próprio arquivo de checklist quando a pasta
    de anexos é a mesma pasta onde ele está). Devolve [] se a pasta não
    existir."""
    if not os.path.isdir(pasta):
        return []
    excluir_abs = {os.path.abspath(c) for c in (excluir or ())}
    return sorted(
        os.path.join(pasta, nome)
        for nome in os.listdir(pasta)
        if os.path.isfile(os.path.join(pasta, nome))
        and not nome.startswith(".")
        and os.path.abspath(os.path.join(pasta, nome)) not in excluir_abs
    )


def encontrar_checklist_na_pasta(pasta: str) -> str:
    """Procura, dentro de `pasta`, o arquivo de checklist: um .docx cujo
    nome (ou, na falta disso, texto) contenha o código de um fluxo
    conhecido (ex.: 'FQ415-031'). Arquivos com 'checklist' no nome são
    testados primeiro, por serem o candidato mais provável. Devolve o
    caminho do arquivo encontrado, ou None se nenhum .docx da pasta bater
    com um fluxo conhecido."""
    candidatos = [c for c in _listar_arquivos_pasta(pasta) if c.lower().endswith(".docx")]
    candidatos.sort(key=lambda c: ("checklist" not in os.path.basename(c).lower(), c))

    for caminho in candidatos:
        if identificar_fluxo(caminho):
            return caminho

    return None


def localizar_anexos(pasta_anexos: str, candidatos: list, excluir: set = None) -> tuple:
    """
    Varre `pasta_anexos` e tenta casar cada arquivo com um dos documentos
    esperados pelo fluxo (`candidatos`, ver fluxos.documentos_do_fluxo),
    reaproveitando a mesma detecção de tipo/identificação de documento do
    main.py (só que restrita aos documentos deste fluxo). Se `excluir` for
    passado, os caminhos nele (por exemplo, o próprio arquivo de checklist,
    quando ele está na mesma pasta dos anexos) são ignorados na varredura.

    Devolve (encontrados, nao_identificados):
      - encontrados: {nome_documento: [(caminho, tipo), ...]}
      - nao_identificados: [caminho, ...] com os arquivos da pasta que não
        bateram com nenhum documento esperado por este fluxo.
    """
    encontrados = {}
    nao_identificados = []

    for caminho in _listar_arquivos_pasta(pasta_anexos, excluir=excluir):
        tipo = automacao_main.detectar_tipo(caminho)
        if tipo is None:
            nao_identificados.append(caminho)
            continue

        texto = automacao_main.extrair_texto_para_identificacao(caminho, tipo)
        documento = doc_types.identificar_documento(caminho, texto, candidatos)

        if documento is None:
            nao_identificados.append(caminho)
            continue

        encontrados.setdefault(documento, []).append((caminho, tipo))

    return encontrados, nao_identificados


def processar_documento_anexo(caminho: str, tipo: str, campos: list) -> None:
    """Chama pdfs.py (PDF) ou tabelas.py (planilha) para extrair, do arquivo
    em `caminho`, os `campos` do checklist. O item especial de conferência
    de assinaturas (fluxos.eh_item_assinatura) não é passado como campo de
    busca; é conferido à parte, usando pdfs.verificar_assinatura."""
    campos_reais = [c for c in campos if not fluxos.eh_item_assinatura(c)]
    precisa_conferir_assinatura = any(fluxos.eh_item_assinatura(c) for c in campos)

    if tipo == "pdf":
        try:
            resultado = pdfs.processar_pdf(caminho, campos_reais)
        except Exception as erro:
            print(f"    [ERRO] Não foi possível abrir/ler '{caminho}': {erro}")
            return

        for campo in campos_reais:
            valor = resultado.get(campo)
            if valor:
                print(f"    [OK] {campo}: {valor}")
            else:
                print(f"    [PENDENTE] {campo}: não encontrado no documento")

        if precisa_conferir_assinatura:
            assinatura = resultado.get("_assinatura", {})
            print()
            if assinatura.get("assinado"):
                sistema = assinatura.get("sistema") or "não identificado"
                print(f"    [ASSINATURA] Assinatura: documento assinado (sistema: {sistema})")
                for s in assinatura.get("signatarios", []):
                    print(f"         - {s['nome']} ({s['cargo']}) em {s['data']} {s['hora']}")
            else:
                print("    [ASSINATURA] Assinatura: nenhum indício de assinatura eletrônica encontrado")

    else:  # xlsx
        try:
            resultado = tabelas.processar_xlsx(caminho, campos_reais)
        except Exception as erro:
            print(f"    [ERRO] Não foi possível abrir '{caminho}': {erro}")
            return

        if resultado is None:
            print(f"    [ERRO] Não foi encontrada nenhuma tabela reconhecível em '{caminho}'.")
            return

        for campo in campos_reais:
            valores = resultado.get(campo)
            if valores is None:
                print(f"    [PENDENTE] {campo}: não encontrado na planilha")
            elif len(valores) == 1:
                print(f"    [OK] {campo}: {valores[0]}")
            else:
                print(f"    [OK] {campo} ({len(valores)} linhas): {valores}")


def executar_checklist(id_fluxo: str, pasta_anexos: str = PASTA_ANEXOS_PADRAO,
                        excluir: set = None) -> None:
    """
    Identifica o fluxo `id_fluxo`, localiza automaticamente (dentro de
    `pasta_anexos`) os arquivos que correspondem a cada documento do
    checklist e chama pdfs.py/tabelas.py para extrair os campos de cada
    um — em vez de só listar os campos esperados. `excluir` permite pular,
    na varredura, arquivos que não são anexos (ex.: o próprio arquivo de
    checklist, quando `pasta_anexos` é a pasta onde ele está).
    """
    dados = fluxos.FLUXOS[id_fluxo]
    documentos_do_fluxo = fluxos.documentos_do_fluxo(id_fluxo)

    print(f"Fluxo identificado: {id_fluxo} - {dados['nome']}\n")
    print(f"Documentos atrelados a este checklist ({len(documentos_do_fluxo)}):\n")

    encontrados, nao_identificados = localizar_anexos(pasta_anexos, documentos_do_fluxo, excluir=excluir)

    for nome_documento in documentos_do_fluxo:
        campos = fluxos.campos_do_documento(id_fluxo, nome_documento)
        print(f"- {nome_documento}")

        arquivos = encontrados.get(nome_documento)
        if not arquivos:
            print(f"    [ERRO] Nenhum arquivo encontrado em '{pasta_anexos}' para este documento.")
            print()
            continue

        for caminho, tipo in arquivos:
            print(f"    Anexo: {os.path.basename(caminho)} "
                  f"({'PDF' if tipo == 'pdf' else 'planilha'})")
            processar_documento_anexo(caminho, tipo, campos)
        print()

    if nao_identificados:
        print(f"[AVISO] Arquivo(s) em '{pasta_anexos}' não identificado(s) como nenhum documento "
              f"deste checklist ({len(nao_identificados)}):")
        for caminho in nao_identificados:
            print(f"  - {os.path.basename(caminho)}")
        print()


def main():
    if len(sys.argv) not in (2, 3):
        print("Uso: python checklist.py caminho/da/pasta")
        print("     python checklist.py caminho/do/checklist.docx [pasta/de/anexos]")
        print("     python checklist.py FQ415-031 [pasta/de/anexos]")
        print(f"     (se omitida, a pasta de anexos padrão é '{PASTA_ANEXOS_PADRAO}')")
        sys.exit(1)

    entrada = sys.argv[1]
    excluir = None

    # Caso 1: `entrada` é uma PASTA -> acha o checklist sozinho dentro dela
    # e usa essa MESMA pasta como pasta de anexos (ignorando, na varredura
    # de anexos, o próprio arquivo de checklist).
    if os.path.isdir(entrada):
        caminho_checklist = encontrar_checklist_na_pasta(entrada)
        if caminho_checklist is None:
            fluxos_conhecidos = ", ".join(fid for fid, _ in fluxos.listar_fluxos())
            print(f"[ERRO] Não encontrei, dentro de '{entrada}', nenhum .docx de checklist "
                  f"que corresponda a um fluxo conhecido. Fluxos reconhecidos: {fluxos_conhecidos}.")
            sys.exit(1)

        id_fluxo = identificar_fluxo(caminho_checklist)
        pasta_anexos = entrada
        excluir = {caminho_checklist}

        print(f"Checklist encontrado: {os.path.basename(caminho_checklist)}\n")

        executar_checklist(id_fluxo, pasta_anexos, excluir=excluir)
        return

    # Caso 2 (compatibilidade): `entrada` é o caminho do .docx do checklist,
    # ou o código do fluxo direto (ex.: 'FQ415-031'), com pasta de anexos
    # separada e opcional.
    pasta_anexos = sys.argv[2] if len(sys.argv) == 3 else PASTA_ANEXOS_PADRAO

    codigo_direto = _normalizar_codigo(entrada) if PADRAO_CODIGO_FLUXO.fullmatch(entrada.strip()) else None

    if codigo_direto in fluxos.FLUXOS:
        id_fluxo = codigo_direto
    else:
        if not os.path.isfile(entrada):
            print(f"[ERRO] Arquivo ou pasta não encontrado(a): '{entrada}'")
            sys.exit(1)
        id_fluxo = identificar_fluxo(entrada)
        excluir = {entrada}

    if id_fluxo is None:
        fluxos_conhecidos = ", ".join(fid for fid, _ in fluxos.listar_fluxos())
        print(f"[ERRO] Não foi possível identificar a qual fluxo o checklist "
              f"'{entrada}' pertence. Fluxos reconhecidos: {fluxos_conhecidos}.")
        sys.exit(1)

    executar_checklist(id_fluxo, pasta_anexos, excluir=excluir)


if __name__ == "__main__":
    main()