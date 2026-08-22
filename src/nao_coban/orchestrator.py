"""
Orquestra o processamento de uma pasta de documentos do fluxo "Não COBAN"
(Cadastro/Atualização comum de fornecedor).
Independente do fluxo COBAN (src.coban.orchestrator), com seu próprio
catálogo de documentos e validadores.

Versão standalone/desktop: SEM fallback de IA. Quando a extração de um
documento é fraca demais pra confiar, o documento é marcado como
`precisa_verificacao_humana=True`. Uma pessoa precisa olhar o arquivo
original.
"""

import concurrent.futures
import logging
import os
import re
import time
import unicodedata

from src.nao_coban.config import TIPOS_FORNECEDOR
from src.nao_coban.validators import (
    detectar_tipo_pessoa_fq064,
    validate_document_nao_coban,
)
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
    nome_empresarial_confere,
    nomes_correspondem,
)
from src.shared.zip_utils import extract_zip, find_os_folders, os_id_from_folder

logger = logging.getLogger(__name__)


def process_os_folder_nao_coban(
    folder_path: str, os_id: str, tipo_fornecedor: str
) -> OsReport:
    """tipo_fornecedor: uma das chaves de TIPOS_FORNECEDOR (ex.:
    "FORNECEDOR_CONTRATADO", "CONCESSIONARIA_SERVICO_PUBLICO", ...).

    Para os tipos que insumos dependem de Pessoa Jurídica x Pessoa Física
    (FORNECEDOR_CONTRATADO e OBRIGACOES_JUDICIAIS), o tipo de pessoa é
    detectado automaticamente a partir do próprio FQ415-064. O analista só
    escolhe o tipo de fornecedor.
    """
    inicio = time.perf_counter()
    tipo_cfg = TIPOS_FORNECEDOR[tipo_fornecedor]
    depende_de_pessoa = "insumos" not in tipo_cfg

    # Pra classificação, usa a união PJ+PF quando depende de pessoa.
    if depende_de_pessoa:
        doc_types_classificacao = {**tipo_cfg["insumos_pj"], **tipo_cfg["insumos_pf"]}
    else:
        doc_types_classificacao = tipo_cfg["insumos"]

    report = OsReport(os_id=os_id)
    print(
        f"\n{'=' * 70}\n### OS: {os_id}  |  tipo: {tipo_fornecedor}  |  pasta: {folder_path}\n{'=' * 70}"
    )
    found_types = set()
    extraction_texts = {}

    files = sorted(
        f
        for f in os.listdir(folder_path)
        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS
    )

    def _processar_arquivo(fname: str):
        arquivo_inicio = time.perf_counter()
        fpath = os.path.join(folder_path, fname)

        tipo_pelo_nome = classify_by_filename(fname, doc_types_classificacao)
        cfg_pelo_nome = (
            doc_types_classificacao.get(tipo_pelo_nome, {}) if tipo_pelo_nome else {}
        )

        if cfg_pelo_nome.get("fora_de_escopo"):
            # Documento fora de escopo (Projeto Básico, Especificações
            # Técnicas...): o conteúdo dele nunca é usado
            # sempre válido.
            extraction = ExtractionResult(
                filename=fname, text="", method="skipped", pages=0
            )
            doc_type = tipo_pelo_nome
            validation = validate_document_nao_coban(doc_type, "", fname)
        else:
            if tipo_pelo_nome == "CONTRATO_SOCIAL":
                # Contrato Social/Ata/Alteração: CNPJ e nome empresarial
                # quase sempre já estão na 1ª página; extract_progressivo
                # faz OCR página a página e para assim que achar os dois,
                # em vez de sempre pagar o teto de OCR_MAX_PAGES páginas de
                # uma vez. Se a 1ª página não bastar, continua normalmente
                # até o mesmo teto de sempre.
                def _contrato_social_tem_dados_suficientes(
                    texto_acumulado, _fname=fname
                ):
                    v = validate_document_nao_coban(
                        "CONTRATO_SOCIAL", texto_acumulado, _fname
                    )
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

            doc_type = classify_document(
                fname, extraction.text, doc_types_classificacao
            )
            validation: DocumentValidation = validate_document_nao_coban(
                doc_type, extraction.text, fname
            )

            if not validation.valido and extraction.method == "ocr":
                print(
                    f"{fname}: 1ª passada de OCR insuficiente, tentando 2ª passada..."
                )
                extraction2 = extract(fpath, two_pass=True)
                doc_type2 = classify_document(
                    fname, extraction2.text, doc_types_classificacao
                )
                validation2 = validate_document_nao_coban(
                    doc_type2, extraction2.text, fname
                )
                if validation2.valido or len(validation2.pendencias) < len(
                    validation.pendencias
                ):
                    extraction, doc_type, validation = (
                        extraction2,
                        doc_type2,
                        validation2,
                    )

        cfg = doc_types_classificacao.get(doc_type, {})
        fora_de_escopo = bool(cfg.get("fora_de_escopo"))

        pendencias = [] if fora_de_escopo else list(validation.pendencias)
        if not fora_de_escopo and extraction.method == "failed" and extraction.error:
            pendencias.insert(
                0,
                f"Falha técnica na leitura do arquivo (OCR/extração não executou): {extraction.error}",
            )

        precisa_verificacao_humana = False
        if not fora_de_escopo:
            min_confidence = cfg.get("min_ocr_confidence", DEFAULT_MIN_OCR_CONFIDENCE)
            precisa_verificacao_humana = extraction.needs_review(min_confidence)
            if precisa_verificacao_humana:
                pendencias.insert(
                    0,
                    "Confiança de leitura insuficiente. Verificação humana necessária.",
                )

        row = DocumentReportRow(
            filename=fname,
            doc_type=doc_type,
            doc_label=cfg.get("label", "Documento não identificado"),
            metodo_leitura=extraction.method,
            confianca_ocr=extraction.ocr_confidence
            or (100.0 if extraction.method == "native_text" else 0.0),
            valido=(
                True
                if fora_de_escopo
                else (validation.valido and not precisa_verificacao_humana)
            ),
            pendencias=pendencias,
            campos=validation.campos,
            erro_extracao=extraction.error if extraction.method == "failed" else None,
            precisa_verificacao_humana=precisa_verificacao_humana,
        )
        arquivo_fim = time.perf_counter()
        print(f"{fname} => {arquivo_fim - arquivo_inicio:.2f}s")
        return fname, extraction.text, fpath, row

    formularios_fname = [
        fname
        for fname in files
        if classify_by_filename(fname, doc_types_classificacao)
        == "FORMULARIO_CADASTRO_FORNECEDOR"
    ]
    max_workers = int(os.environ.get("OCR_MAX_WORKERS", "4")) + len(formularios_fname)
    file_paths = {}
    executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
    try:
        # Dispara a OCR forçada do FQ415-064 em paralelo com o lote principal.
        # Só é consultada se o texto nativo não confirmar a assinatura.
        # extract_ocr_forced_skip_edges pula a primeira/última
        # página, que nunca têm o carimbo de assinatura. dpi=150 (abaixo do
        # padrão de 300) porque aqui só precisamos DETECTAR um carimbo/
        # palavra-chave (D4Sign, gov.br, certificado ICP-Brasil), não ler
        # dado fino resolução menor já basta.
        futures_assinatura = {
            fname: executor.submit(
                extract_ocr_forced_skip_edges,
                os.path.join(folder_path, fname),
                two_pass=False,
                dpi=150,
            )
            for fname in formularios_fname
        }

        resultados = list(executor.map(_processar_arquivo, files))
    finally:
        executor.shutdown(wait=False)

    for fname, texto, fpath, row in resultados:
        found_types.add(row.doc_type)
        extraction_texts[fname] = texto
        file_paths[fname] = fpath
        report.rows.append(row)

    # Se o tipo depende de PJ/PF, decide agora (com o FQ415-064 já lido)
    # qual conjunto de insumos exigir.
    tipo_pessoa = None
    if depende_de_pessoa:
        formulario_row = next(
            (r for r in report.rows if r.doc_type == "FORMULARIO_CADASTRO_FORNECEDOR"),
            None,
        )
        if formulario_row:
            tipo_pessoa = formulario_row.campos.get("tipo_pessoa")
            if tipo_pessoa not in ("PJ", "PF"):
                tipo_pessoa = detectar_tipo_pessoa_fq064(
                    extraction_texts.get(formulario_row.filename, "")
                )
        doc_types_exigidos = (
            tipo_cfg["insumos_pf"] if tipo_pessoa == "PF" else tipo_cfg["insumos_pj"]
        )
        if tipo_pessoa is None:
            report.pendencias_gerais.append(
                "Não foi possível identificar automaticamente se o cadastro é de Pessoa Jurídica ou "
                "Pessoa Física a partir do FQ415-064. Verificar manualmente qual seção do formulário "
                "está preenchida e conferir os insumos correspondentes na TAB415-004."
            )
    else:
        doc_types_exigidos = tipo_cfg["insumos"]

    for doc_type, cfg in doc_types_exigidos.items():
        if cfg["required"] and doc_type not in found_types:
            report.documentos_faltantes.append(cfg["label"])

    # Validação cruzada de CNPJ (ignora DOCUMENTO_FORA_DE_ESCOPO, que nunca
    # tem campos["cnpj"] preenchido de qualquer forma).
    validations = [
        DocumentValidation(
            doc_type=r.doc_type, filename=r.filename, valido=r.valido, campos=r.campos
        )
        for r in report.rows
    ]
    report.pendencias_gerais.extend(cross_validate_cnpj(validations))

    #    Confirmação de que alguém assinou o FQ415-064 (D4Sign,
    #    gov.br ou certificado ICP-Brasil) independente de existir
    #    Contrato Social no pacote ou não. validate_formulario_cadastro_
    #    fornecedor() já deixou essa pendência quando o texto nativo não
    #    confirmou; aqui só tenta resolver com a OCR forçada antes de aceitar.
    # Precisa ser IDÊNTICA à mensagem gerada em
    # validate_formulario_cadastro_fornecedor (src/nao_coban/validators.py);
    # é comparada por igualdade de string logo abaixo pra decidir se remove
    # a pendência; qualquer diferença de texto entre as duas quebra essa
    # remoção silenciosamente.
    pendencia_assinatura = (
        "Não foi possível confirmar assinatura (D4Sign ou certificado digital) no documento."
    )
    # Um pacote pode ter mais de um arquivo classificado como FQ415-064
    # (um modelo em branco reanexado por engano junto com o assinado de
    # verdade) confirma a assinatura em CADA UM, não só no primeiro.
    formulario_rows = [
        r for r in report.rows if r.doc_type == "FORMULARIO_CADASTRO_FORNECEDOR"
    ]
    for formulario_row in formulario_rows:
        nome_assinante = formulario_row.campos.get("nomes_socios", "")
        nomes_assinantes_texto = [
            n
            for n in formulario_row.campos.get("nomes_assinantes_d4sign", "").split(";")
            if n
        ]
        assinatura_confirmada = bool(
            re.search(
                r"d4sign", extraction_texts.get(formulario_row.filename, ""), re.I
            )
            or (
                nome_assinante
                and any(
                    nomes_correspondem(nome_assinante, n)
                    for n in nomes_assinantes_texto
                )
            )
        )
        if assinatura_confirmada:
            formulario_row.campos["metodo_assinatura"] = (
                "D4Sign/certificado (texto nativo)"
            )

        if not assinatura_confirmada:
            future = futures_assinatura.get(formulario_row.filename)
            ocr_inicio = time.perf_counter()
            try:
                texto_ocr = (future.result() or "") if future else ""
            except Exception as exc:
                logger.warning(
                    "OCR forçado de assinatura falhou (%s): %s",
                    formulario_row.filename,
                    exc,
                )
                texto_ocr = ""
            ocr_fim = time.perf_counter()
            print(
                f"OCR forçado ({formulario_row.filename}) => {ocr_fim - ocr_inicio:.2f}s"
            )
            texto_ocr_norm = _normalizar_nome(texto_ocr)
            # _normalizar_nome tira pontuação (inclusive ':')
            texto_ocr_com_dois_pontos = (
                unicodedata.normalize("NFKD", texto_ocr or "")
                .encode("ascii", "ignore")
                .decode("ascii")
            )
            texto_ocr_com_dois_pontos = re.sub(
                r"[^A-Za-z0-9:\s]", " ", texto_ocr_com_dois_pontos
            ).upper()
            partes_nome = nome_assinante.split() if nome_assinante else []
            sobrenome_norm = _normalizar_nome(partes_nome[-1]) if partes_nome else ""

            # D4Sign nativo/gov.br: procura o carimbo (sinal de que
            # alguém assinou).
            trechos = []
            for kw in (
                r"d4sign",
                r"documento assinado digitalmente",
                r"assinado\s+de\s+forma\s+digital\s+por",
                r"assinante",
                r"netlex",
            ):
                # texto_ocr_norm vem de _normalizar_nome(), que sempre
                # devolve tudo em maiúsculas. re.I é obrigatório aqui porque
                # essas palavras-chave estão em minúsculo.
                for m in re.finditer(kw, texto_ocr_norm, re.I):
                    trechos.append(
                        texto_ocr_norm[max(0, m.start() - 90) : m.end() + 90]
                    )
            if trechos and (
                not sobrenome_norm
                or any(sobrenome_norm in trecho for trecho in trechos)
            ):
                assinatura_confirmada = True
                formulario_row.campos["metodo_assinatura"] = "D4Sign/gov.br (OCR)"

            # Certificado ICP-Brasil: "Assinado de forma digital por" tem
            # fonte pequena demais no carimbo, o Tesseract descarta essa
            # linha mesmo em 300 DPI mas o padrão "SOBRENOME:CPF" sobrevive ao OCR de forma confiável.
            if not assinatura_confirmada:
                for m in re.finditer(
                    r"([A-Z]{3,})\s*:\s*\d[\d\s]{8,14}\d", texto_ocr_com_dois_pontos
                ):
                    if not sobrenome_norm or sobrenome_norm in m.group(1):
                        assinatura_confirmada = True
                        formulario_row.campos["metodo_assinatura"] = (
                            "certificado ICP-Brasil (OCR, carimbo NOME:CPF)"
                        )
                        break

        formulario_row.campos["assinatura_confirmada"] = str(assinatura_confirmada)
        if assinatura_confirmada:
            formulario_row.pendencias = [
                p for p in formulario_row.pendencias if p != pendencia_assinatura
            ]
            formulario_row.valido = not formulario_row.pendencias

    #    Quem assina x Contrato Social, só faz sentido pra Pessoa Jurídica
    #    com Contrato Social exigido (TAB415-004: nos tipos de Não
    #    Fornecedor "fixos", quem assina costuma ser o gestor da área
    #    solicitante, não o representante legal do fornecedor).
    if tipo_pessoa == "PJ" or (
        not depende_de_pessoa
        and "CONTRATO_SOCIAL" in doc_types_exigidos
        and tipo_fornecedor == "FORNECEDOR_CONTRATADO"
    ):
        # Uma OS pode ter mais de um documento classificado como
        # CONTRATO_SOCIAL (ex.: Ata de Constituição + Alteração/Consolidação
        # de Estatuto, comum em associações). O nome de quem assina pode
        # estar em só um deles (normalmente a Ata, que lista quem foi eleito;
        # o Estatuto em si costuma ser só regras gerais, sem nomes). Por
        # isso a checagem busca em TODOS, não só no primeiro encontrado.
        contrato_rows = [r for r in report.rows if r.doc_type == "CONTRATO_SOCIAL"]
        contrato_row = contrato_rows[0] if contrato_rows else None
        texto_contrato_combinado = "\n".join(
            extraction_texts.get(r.filename, "") for r in contrato_rows
        )
        procuracao_row = next(
            (r for r in report.rows if r.doc_type == "PROCURACAO"), None
        )
        for formulario_row in formulario_rows:
            if formulario_row and contrato_row:
                nome_assinante = formulario_row.campos.get("nomes_socios", "")
                texto_contrato = texto_contrato_combinado
                assinante_e_socio = bool(nome_assinante) and nome_empresarial_confere(
                    nome_assinante, texto_contrato
                )

                if nome_assinante and not assinante_e_socio:
                    # A extração do Contrato Social pode ter parado cedo por já ter
                    # achado CNPJ/nome empresarial na 1ª página mas o nome
                    # de quem assina (oficial eleito numa Ata de Eleição)
                    # pode só aparecer numa página seguinte. Antes de reprovar,
                    # tenta 1x uma releitura completa (até OCR_MAX_PAGES) dos arquivos de Contrato
                    # Social cujo texto ainda não tem esse nome custo extra
                    # só nesse caso pontual, não no fluxo normal.
                    releu_algum = False
                    for r in contrato_rows:
                        texto_atual = extraction_texts.get(r.filename, "")
                        if nome_empresarial_confere(nome_assinante, texto_atual):
                            continue  # esse já tem o nome, não precisa reler
                        fpath_contrato = file_paths.get(r.filename)
                        if not fpath_contrato:
                            continue
                        releitura_inicio = time.perf_counter()
                        texto_completo = extract(fpath_contrato).text
                        releitura_fim = time.perf_counter()
                        print(
                            f"Releitura completa do Contrato Social ({r.filename}) => {releitura_fim - releitura_inicio:.2f}s"
                        )
                        if texto_completo and len(texto_completo) > len(texto_atual):
                            extraction_texts[r.filename] = texto_completo
                            releu_algum = True
                    if releu_algum:
                        texto_contrato_combinado = "\n".join(
                            extraction_texts.get(r.filename, "") for r in contrato_rows
                        )
                        texto_contrato = texto_contrato_combinado
                        assinante_e_socio = bool(
                            nome_assinante
                        ) and nome_empresarial_confere(nome_assinante, texto_contrato)

                if nome_assinante:
                    # Guarda o resultado dessa conferência num campo próprio
                    # pra app_desktop.py poder montar o painel "Dados da
                    # empresa (justificativa)" sem precisar reconstruir essa
                    # lógica a partir de pendências/valido.
                    formulario_row.campos["nome_socio_conferido"] = (
                        "True" if assinante_e_socio else "False"
                    )

                if nome_assinante and not assinante_e_socio:
                    # Não é sócio antes de reprovar, confere se é um dos
                    # procuradores listados na Procuração (TAB415-004:
                    # a Procuração existe exatamente pra cobrir esse caso).
                    outorgados = []
                    if procuracao_row:
                        outorgados = [
                            n
                            for n in procuracao_row.campos.get("outorgados", "").split(
                                ";"
                            )
                            if n
                        ]
                    assinante_e_procurador = any(
                        nomes_correspondem(nome_assinante, outorgado)
                        for outorgado in outorgados
                    )

                    if assinante_e_procurador:
                        formulario_row.campos["assinatura_via_procuracao"] = "True"
                        formulario_row.campos["procurador_confirmado"] = nome_assinante
                        # Confirmado como procurador. Não é pendência (mas a
                        # pendência de evidência de assinatura se
                        # ainda existir, continua valendo).
                    elif procuracao_row:
                        formulario_row.pendencias.append(
                            f"Quem assina o FQ415-064 ('{nome_assinante}') não foi encontrado no Contrato Social "
                            f"nem entre os outorgados da Procuração no pacote ({procuracao_row.filename}). "
                            "Confirmar manualmente quem é essa pessoa."
                        )
                        formulario_row.valido = False
                    else:
                        formulario_row.pendencias.append(
                            f"Quem assina o FQ415-064 ('{nome_assinante}') não foi encontrado no Contrato Social "
                            "e não há Procuração no pacote que justifique outra pessoa assinar em nome do "
                            "fornecedor. Verificar manualmente."
                        )
                        formulario_row.valido = False
            elif (
                formulario_row
                and not contrato_row
                and "CONTRATO_SOCIAL" in doc_types_exigidos
            ):
                pass  # já reportado em documentos_faltantes

    #    Nome de referência (FQ415-064) x Documento de Identificação
    #    mais confiável que tentar "ler" o nome direto do RG/CNH:
    #    o layout varia muito entre documentos e o OCR embaralha a ordem
    #    dos campos do cartão. O nome da pessoa já é conhecido de forma
    #    confiável pelo campo "Nome:" do próprio FQ415-064. Só confere se
    #    ele aparece no texto do documento de identificação, igual já se
    #    faz pra CNPJ/nome empresarial no Contrato Social.
    nome_referencia_pf = next(
        (
            fr.campos.get("nomes_socios", "")
            for fr in formulario_rows
            if fr.campos.get("nomes_socios")
        ),
        "",
    )
    for doc_id_row in [
        r for r in report.rows if r.doc_type == "DOCUMENTO_IDENTIFICACAO"
    ]:
        if not nome_referencia_pf:
            continue  # sem nome de referência (FQ415-064 não encontrado/sem nome) nada pra conferir
        texto_doc_id = extraction_texts.get(doc_id_row.filename, "")
        confere = nome_empresarial_confere(nome_referencia_pf, texto_doc_id)
        doc_id_row.campos["nome"] = nome_referencia_pf if confere else ""
        doc_id_row.campos["nome_confere_fq064"] = str(confere)
        if not confere:
            doc_id_row.pendencias.append(
                f"Nome de referência do FQ415-064 ('{nome_referencia_pf}') não localizado neste documento "
                "de identificação. Verificar manualmente."
            )
            doc_id_row.valido = False

    fim = time.perf_counter()
    print(
        f"\nOS {os_id} ({tipo_fornecedor}) processada em {fim - inicio:.2f} segundos\n"
    )
    report.tempo_processamento_segundos = fim - inicio
    return report


def process_zip_nao_coban(
    zip_path: str,
    tipo_fornecedor: str,
    fallback_os_id: str = "0",
    workdir: str = "/tmp",
) -> list[OsReport]:
    dest_dir = os.path.join(workdir, f"extract_naocoban_{os.path.basename(zip_path)}")
    os.makedirs(dest_dir, exist_ok=True)
    extract_zip(zip_path, dest_dir)

    folders = find_os_folders(dest_dir)
    reports = []
    for folder in folders:
        os_id = os_id_from_folder(folder, fallback_os_id)
        reports.append(process_os_folder_nao_coban(folder, os_id, tipo_fornecedor))
    return reports
