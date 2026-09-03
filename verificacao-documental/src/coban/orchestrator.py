"""
Orquestra o processamento de uma pasta de documentos do fluxo COBAN
(Credenciamento de Correspondente Bancário), juntando classificação,
extração (OCR em camadas) e validação.

Versão standalone/desktop: SEM fallback de IA. Quando a extração de um
documento é fraca demais pra confiar (comprovante bancário borrado),
o documento é marcado como `precisa_verificacao_humana=True`.
Uma pessoa precisa olhar o arquivo original.
"""

import concurrent.futures
import logging
import os
import re
import time

from src.coban.config import DOCUMENT_TYPES
from src.coban.validators import validate_document
from src.shared.classifier import classify_by_filename, classify_document
from src.shared.constants import DEFAULT_MIN_OCR_CONFIDENCE, SUPPORTED_EXTENSIONS
from src.shared.extractor import (
    ExtractionResult,
    extract,
    extract_ocr_forced_skip_edges,
    extract_progressivo,
)
from src.shared.report import DocumentReportRow, OsReport
from src.shared.validation import (
    DocumentValidation,
    _normalizar_nome,
    cross_validate_cnpj,
    find_all_cnpjs,
    find_cnpj,
    nome_empresarial_confere,
    nomes_correspondem,
)
from src.shared.zip_utils import extract_zip, find_os_folders, os_id_from_folder

logger = logging.getLogger(__name__)


def process_os_folder(folder_path: str, os_id: str) -> OsReport:
    inicio = time.perf_counter()

    report = OsReport(os_id=os_id)
    print(f"\n{'=' * 70}\n### OS: {os_id}  |  pasta: {folder_path}\n{'=' * 70}")
    found_types = set()
    extraction_texts = {}
    file_paths = {}

    files = sorted(
        f
        for f in os.listdir(folder_path)
        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
    )

    def _processar_arquivo(fname: str):
        """Processa 1 arquivo (extração + classificação + validação).
        Roda em thread separada."""
        arquivo_inicio = time.perf_counter()
        fpath = os.path.join(folder_path, fname)

        # Contrato Social/Ata/Alteração escaneado costuma ser o arquivo mais
        # pesado da OS (CNPJ e nome empresarial praticamente sempre estão na
        # 1ª página): extract_progressivo faz OCR página a página e para
        # assim que achar os dois, em vez de sempre pagar o teto de
        # OCR_MAX_PAGES páginas de uma vez
        tipo_pelo_nome = classify_by_filename(fname, DOCUMENT_TYPES)
        if tipo_pelo_nome == "CONTRATO_SOCIAL":

            def _contrato_social_tem_dados_suficientes(texto_acumulado, _fname=fname):
                v = validate_document("CONTRATO_SOCIAL", texto_acumulado, _fname)
                return v.valido and bool(v.campos.get("nome_empresarial"))

            extraction: ExtractionResult = extract_progressivo(
                fpath, _contrato_social_tem_dados_suficientes
            )
        else:
            extraction = extract(fpath)

        if extraction.method == "failed" and extraction.error:
            logger.error(
                "Falha na extração de '%s' (%s): %s", fname, os_id, extraction.error
            )

        doc_type = classify_document(fname, extraction.text, DOCUMENT_TYPES)
        validation: DocumentValidation = validate_document(
            doc_type, extraction.text, fname
        )

        if not validation.valido and extraction.method == "ocr":
            print(f"{fname}: 1ª passada de OCR insuficiente, tentando 2ª passada...")
            extraction2 = extract(fpath, two_pass=True)
            doc_type2 = classify_document(fname, extraction2.text, DOCUMENT_TYPES)
            validation2 = validate_document(doc_type2, extraction2.text, fname)
            if validation2.valido or len(validation2.pendencias) < len(
                validation.pendencias
            ):
                extraction, doc_type, validation = extraction2, doc_type2, validation2

        cfg = DOCUMENT_TYPES.get(doc_type, {})
        pendencias = list(validation.pendencias)
        if extraction.method == "failed" and extraction.error:
            pendencias.insert(
                0,
                f"Falha técnica na leitura do arquivo (OCR/extração não executou): {extraction.error}",
            )

        min_confidence = cfg.get("min_ocr_confidence", DEFAULT_MIN_OCR_CONFIDENCE)
        precisa_verificacao_humana = extraction.needs_review(min_confidence)
        if precisa_verificacao_humana:
            pendencias.insert(
                0, "Confiança de leitura insuficiente. Verificação humana necessária."
            )

        row = DocumentReportRow(
            filename=fname,
            doc_type=doc_type,
            doc_label=cfg.get("label", "Documento não identificado"),
            metodo_leitura=extraction.method,
            confianca_ocr=extraction.ocr_confidence
            or (100.0 if extraction.method == "native_text" else 0.0),
            valido=validation.valido and not precisa_verificacao_humana,
            pendencias=pendencias,
            campos=validation.campos,
            erro_extracao=extraction.error if extraction.method == "failed" else None,
            precisa_verificacao_humana=precisa_verificacao_humana,
        )

        arquivo_fim = time.perf_counter()
        print(f"{fname} => {arquivo_fim - arquivo_inicio:.2f}s")

        return fname, extraction.text, fpath, row

    max_workers = int(os.environ.get("OCR_MAX_WORKERS", "4"))
    executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
    try:
        futures_assinatura = {
            fname: executor.submit(
                extract_ocr_forced_skip_edges,
                os.path.join(folder_path, fname),
                two_pass=True,
            )
            for fname in files
            if re.search(r"d4sign", fname, re.I)
        }

        resultados = list(executor.map(_processar_arquivo, files))
    finally:
        executor.shutdown(wait=False)

    for fname, texto, fpath, row in resultados:
        found_types.add(row.doc_type)
        extraction_texts[fname] = texto
        file_paths[fname] = fpath
        report.rows.append(row)

    # documentos obrigatórios ausentes no pacote
    for doc_type, cfg in DOCUMENT_TYPES.items():
        if cfg["required"] and doc_type not in found_types:
            report.documentos_faltantes.append(cfg["label"])

    # validação cruzada (mesmo CNPJ em todos os documentos)
    validations = [
        DocumentValidation(
            doc_type=r.doc_type, filename=r.filename, valido=r.valido, campos=r.campos
        )
        for r in report.rows
    ]
    report.pendencias_gerais.extend(cross_validate_cnpj(validations))

    # Cartão CNPJ como documento-base: quando o Contrato Social/
    #    Alteração Contratual não traz CNPJ próprio, confere se o nome
    #    empresarial do Cartão CNPJ aparece no texto do contrato, em vez de
    #    reprovar por "CNPJ não localizado".
    cartao_row = next(
        (r for r in report.rows if r.doc_type in ("CARTAO_CNPJ", "CNPJ")), None
    )
    nome_base = (cartao_row.campos.get("nome_empresarial") if cartao_row else "") or ""
    cnpj_base = (cartao_row.campos.get("cnpj") if cartao_row else "") or ""

    for row in report.rows:
        if row.doc_type != "CONTRATO_SOCIAL":
            continue

        texto_contrato = extraction_texts.get(row.filename, "")
        row.pendencias = [p for p in row.pendencias if "CNPJ não localizado" not in p]

        cnpjs_no_contrato = find_all_cnpjs(texto_contrato)

        if cnpj_base and cnpj_base in cnpjs_no_contrato:
            # CNPJ da empresa (Cartão CNPJ) realmente aparece no contrato.
            # confirma e substitui qualquer CNPJ "errado" (do cartório)
            # que o validador isolado do Contrato Social tenha pego antes.
            row.campos["cnpj"] = cnpj_base
            row.campos["nome_empresarial"] = nome_base or row.campos.get(
                "nome_empresarial", ""
            )
            row.valido = not row.pendencias
            continue

        # CNPJ da empresa não apareceu (ausente de verdade, ou só apareceu o
        # de outra parte, tipo cartório) -> cai pra conferência por nome.
        if not cartao_row or not nome_base:
            row.pendencias.append(
                "CNPJ não localizado no Contrato Social/CCMEI e não há Cartão CNPJ no pacote para conferir pelo nome empresarial."
            )
            row.valido = False
        elif nome_empresarial_confere(nome_base, texto_contrato):
            row.campos["nome_empresarial_conferido"] = "True"
            row.campos["nome_empresarial_referencia"] = nome_base
            row.campos["nome_empresarial"] = nome_base
            row.campos["cnpj"] = cnpj_base
            row.valido = not row.pendencias
        else:
            row.pendencias.append(
                f"Nome empresarial do Contrato Social não confere com o Cartão CNPJ ('{nome_base}'). Verificar manualmente."
            )
            row.campos["nome_empresarial_conferido"] = "False"
            row.campos["nome_empresarial_referencia"] = nome_base
            row.valido = False

    #    O(s) sócio(s) listado(s) no próprio Formulário de Credenciamento
    #    ("Nome Sócio 1:", "Nome Sócio 2:"...) precisa(m) aparecer no
    #    Contrato Social/CCMEI. Essa é a checagem principal, e usa um
    #    campo direto e confiável do formulário.
    #    A confirmação de QUEM assinou (D4Sign ou gov.br)
    #    vira um sinal complementar, não bloqueia a validação sozinha.

    formulario_row = next(
        (r for r in report.rows if r.doc_type == "FORMULARIO_CREDENCIAMENTO"), None
    )
    contrato_row = next(
        (r for r in report.rows if r.doc_type == "CONTRATO_SOCIAL"), None
    )
    texto_contrato = (
        extraction_texts.get(contrato_row.filename, "") if contrato_row else ""
    )

    if formulario_row:
        nomes_socios = [
            n for n in formulario_row.campos.get("nomes_socios", "").split(";") if n
        ]
        nomes_assinantes_d4sign = [
            n
            for n in formulario_row.campos.get("nomes_assinantes_d4sign", "").split(";")
            if n
        ]

        if not nomes_socios:
            pass  # já reportado como pendência em validate_formulario_credenciamento
        elif not contrato_row:
            formulario_row.pendencias.append(
                "Contrato Social/CCMEI não encontrado no pacote; não foi possível "
                "conferir se o(s) sócio(s) do Formulário consta(m) nele."
            )
            formulario_row.campos["nome_socio_conferido"] = "NAO_VERIFICAVEL"
        else:
            socios_confirmados = [
                n for n in nomes_socios if nome_empresarial_confere(n, texto_contrato)
            ]

            formulario_row.campos["nome_socio_conferido"] = str(
                bool(socios_confirmados)
            )

            if not socios_confirmados:
                formulario_row.pendencias.append(
                    f"Nenhum dos sócios listados no Formulário ({', '.join(nomes_socios)}) "
                    "foi encontrado no Contrato Social/CCMEI. Verificar manualmente."
                )
                formulario_row.valido = False
            else:
                # Sinal complementar: alguma evidência de que ESSE sócio
                # (não só alguém da BBTS) realmente assinou o documento
                # via D4Sign nativo (nome bate) ou via gov.br. Se o texto
                # nativo não achou nada, tenta de novo com OCR forçado antes
                # de desistir (custo extra só nesse caso pontual).
                socio_assinante = next(
                    (
                        socio
                        for socio in socios_confirmados
                        for assinante in nomes_assinantes_d4sign
                        if nomes_correspondem(socio, assinante)
                    ),
                    None,
                )

                assinatura_confirmada = socio_assinante is not None
                if assinatura_confirmada:
                    formulario_row.campos["metodo_assinatura_socio"] = "D4Sign"

                if not assinatura_confirmada:
                    caminho_formulario = file_paths.get(formulario_row.filename)

                    if caminho_formulario:
                        ocr_inicio = time.perf_counter()
                        future = futures_assinatura.get(formulario_row.filename)
                        if future is not None:
                            try:
                                texto_ocr = future.result()
                            except Exception as exc:
                                logger.warning(
                                    "OCR forçado de assinatura falhou (%s): %s",
                                    formulario_row.filename,
                                    exc,
                                )
                                texto_ocr = extract_ocr_forced_skip_edges(
                                    caminho_formulario, two_pass=True
                                )
                        else:
                            texto_ocr = extract_ocr_forced_skip_edges(
                                caminho_formulario, two_pass=True
                            )

                        ocr_fim = time.perf_counter()
                        print(
                            f"OCR forçado ({formulario_row.filename}) => {ocr_fim - ocr_inicio:.2f}s"
                        )

                        texto_ocr_norm = _normalizar_nome(texto_ocr or "")

                        # O nome do assinante pode vir ANTES ou DEPOIS de "DOCUMENTO
                        # ASSINADO DIGITALMENTE". Por isso a janela é larga pros dois lados;
                        # finditer cobre PDFs com mais de um carimbo.
                        janela_antes = 15
                        janela_depois = 90
                        trechos_assinatura = [
                            texto_ocr_norm[
                                max(0, m.start() - janela_antes) : m.end()
                                + janela_depois
                            ]
                            for m in re.finditer(
                                r"DOCUMENTO ASSINADO DIGITALMENTE", texto_ocr_norm
                            )
                        ]

                        assinatura_confirmada = False
                        for socio in socios_confirmados:
                            socio_norm = _normalizar_nome(socio)
                            if socio_norm and any(
                                socio_norm in trecho for trecho in trechos_assinatura
                            ):
                                assinatura_confirmada = True
                                socio_assinante = socio
                                break

                        if assinatura_confirmada:
                            formulario_row.campos["metodo_assinatura_socio"] = "gov.br"

                formulario_row.campos.setdefault(
                    "metodo_assinatura_socio", "NAO_CONFIRMADO"
                )
                formulario_row.campos["socio_assinante"] = socio_assinante or ""

                if not assinatura_confirmada:
                    # TAB415-004: uma Procuração pode autorizar
                    # alguém que não é sócio a assinar em nome do
                    # fornecedor. Se existir uma Procuração no pacote,
                    # isso é motivo pra assinatura não bater com
                    # o sócio, não uma pendência real.
                    procuracao_row = next(
                        (r for r in report.rows if r.doc_type == "PROCURACAO"), None
                    )
                    if procuracao_row:
                        formulario_row.campos["assinatura_via_procuracao"] = "True"
                        formulario_row.pendencias.append(
                            f"Assinatura não corresponde ao sócio '{socios_confirmados}', mas há uma "
                            f"Procuração no pacote ({procuracao_row.filename}); confirmar se autoriza "
                            "quem assinou o Formulário."
                        )
                    else:
                        formulario_row.pendencias.append(
                            f"Sócio '{socios_confirmados}' consta no Contrato Social, mas não foi possível "
                            "confirmar evidência de assinatura dele (D4Sign nativo ou gov.br, mesmo com OCR) "
                            "no Formulário, e não há Procuração no pacote que justifique outra pessoa "
                            "assinar em nome dele. Verificar manualmente."
                        )
                        formulario_row.valido = False

    #    Optante pelo Simples Nacional precisa bater entre o Parecer COBAN
    #    e a Consulta Simples Nacional; os dois têm que concordar, seja
    #    "optante" ou "não optante".
    parecer_row = next((r for r in report.rows if r.doc_type == "PARECER_COBAN"), None)
    simples_row = next(
        (r for r in report.rows if r.doc_type == "SIMPLES_NACIONAL"), None
    )

    if parecer_row and simples_row:
        optante_parecer = parecer_row.campos.get("optante_simples")
        optante_simples_doc = simples_row.campos.get("optante_simples")
        if (
            optante_parecer is not None
            and optante_simples_doc is not None
            and optante_parecer != optante_simples_doc
        ):
            msg = (
                f"Divergência sobre opção pelo Simples Nacional: Parecer COBAN diz "
                f"'{optante_parecer}', Consulta Simples Nacional diz '{optante_simples_doc}'. "
                "Verificar manualmente."
            )
            parecer_row.pendencias.append(msg)
            simples_row.pendencias.append(msg)
            parecer_row.valido = False
            simples_row.valido = False

    for row in report.rows:
        # PARECER_COBAN é um documento-modelo genérico (regras de retenção
        # fiscal) e COMPROVANTE_BANCARIO só precisa mostrar dados da conta

        if (
            row.doc_type
            in ("CARTAO_CNPJ", "CNPJ", "PARECER_COBAN", "COMPROVANTE_BANCARIO")
            or not cartao_row
        ):
            continue
        texto_doc = extraction_texts.get(row.filename, "")
        cnpj_bate = bool(cnpj_base) and find_cnpj(texto_doc) == cnpj_base
        nome_bate = bool(nome_base) and nome_empresarial_confere(nome_base, texto_doc)
        if not cnpj_bate and not nome_bate:
            row.pendencias.append(
                "Não foi possível confirmar o CNPJ nem a razão social da empresa "
                "neste documento. Verificar manualmente."
            )
            row.valido = False

    fim = time.perf_counter()

    print(f"\nOS {os_id} processada em {fim - inicio:.2f} segundos\n")

    report.tempo_processamento_segundos = fim - inicio
    return report


def process_zip(
    zip_path: str, fallback_os_id: str = "0", workdir: str = "/tmp"
) -> list[OsReport]:
    """Descompacta o zip (que pode conter documentos soltos ou várias
    subpastas) e processa cada pasta encontrada."""
    dest_dir = os.path.join(workdir, f"extract_{os.path.basename(zip_path)}")
    os.makedirs(dest_dir, exist_ok=True)
    extract_zip(zip_path, dest_dir)

    folders = find_os_folders(dest_dir)
    reports = []
    for folder in folders:
        os_id = os_id_from_folder(folder, fallback_os_id)
        reports.append(process_os_folder(folder, os_id))
    return reports
