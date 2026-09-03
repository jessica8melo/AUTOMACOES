"""
Extração de texto em camadas:

  Camada 1 - pdfplumber: tenta ler o texto nativo do PDF (rápido e
             100% de confiança quando funciona).
  Camada 2 - OCR (pdf2image + pytesseract): usada quando a camada 1 não
             retorna texto (PDF é imagem escaneada) ou quando o texto vem
             visivelmente corrompido. Cobre CNPJ, Consulta
             Optantes, etc., que no material de amostra vêm sempre como
             PDF escaneado.
  Camada 3 - Imagem direta (JPG/PNG): vai direto para o OCR.

O resultado sempre inclui uma confiança estimada (0-100), usada para decidir
se o documento precisa de verificação humana (ver ExtractionResult.needs_review).

Usada pelos dois fluxos (COBAN e Não COBAN); não conhece nenhum dos dois,
só extrai texto.
"""

import os
import re
from dataclasses import dataclass
from typing import Optional

import pdfplumber
import pytesseract
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image, ImageFilter, ImageOps

from src.shared.constants import DEFAULT_MIN_OCR_CONFIDENCE

OCR_LANG = os.environ.get("OCR_LANG", "por")
OCR_DPI = int(os.environ.get("OCR_DPI", "300"))
OCR_PSM = os.environ.get("OCR_PSM", "6")
OCR_BIN_THRESHOLD = int(os.environ.get("OCR_BIN_THRESHOLD", "180"))
OCR_RETRY_THRESHOLD = float(os.environ.get("OCR_RETRY_THRESHOLD", "75"))
OCR_MAX_PAGES = int(os.environ.get("OCR_MAX_PAGES", "5"))

# A partir da v4, o Tesseract usa OpenMP internamente e tenta usar todos os
# núcleos da máquina em CADA chamada. Como o app já roda vários arquivos em
# paralelo (ThreadPoolExecutor, OCR_MAX_WORKERS threads), sem esse limite
# cada thread dispara um tesseract.exe brigando pelos mesmos núcleos,
# deixando o processamento em lote mais lento do que sequencial. Limitar a
# 1 núcleo por chamada deixa o ThreadPoolExecutor como único responsável por
# dividir os núcleos disponíveis. setdefault: não sobrescreve se já tiver
# sido definido explicitamente (execução via linha de comando).
os.environ.setdefault("OMP_THREAD_LIMIT", "1")


def _tesseract_cmd():
    return os.environ.get("TESSERACT_CMD")


def _poppler_path():
    return os.environ.get("POPPLER_PATH")


MIN_NATIVE_TEXT_CHARS = 30  # abaixo disso, "sem texto nativo útil"


@dataclass
class ExtractionResult:
    filename: str
    text: str = ""
    method: str = "none"  # "native_text" | "ocr" | "failed"
    ocr_confidence: Optional[float] = None
    pages: int = 0
    error: Optional[str] = None

    @property
    def needs_ai_fallback(self) -> bool:
        """Sinaliza quando o resultado é fraco demais para confiar,
        usando o limiar GLOBAL (src.shared.constants.DEFAULT_MIN_OCR_CONFIDENCE).
        """
        return self.needs_review(_threshold())

    def needs_review(self, min_confidence: float) -> bool:
        if self.method == "failed":
            return True
        if self.method == "native_text":
            return False  # texto nativo é confiável por definição
        if self.method == "ocr":
            return (self.ocr_confidence or 0) < 100 and (
                self.ocr_confidence or 0
            ) < min_confidence
        return True


def _threshold():
    return DEFAULT_MIN_OCR_CONFIDENCE


def _extract_native_text(pdf_path: str) -> str:
    chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            chunks.append(t)
    return "\n".join(chunks).strip()


def _preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """Pré-processamento leve pra melhorar leitura de documentos
    escaneados/fotografados (scanner de má qualidade, foto...). Só usa Pillow:
      1) escala de cinza: remove ruído de cor/JPEG;
      2) autocontraste: realça a diferença entre texto e fundo;
      3) filtro de mediana: reduz "grão" de scanner/foto sem borrar bordas;
      4) binarização (preto/branco).
    """
    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray, cutoff=1)
    gray = gray.filter(ImageFilter.MedianFilter(size=3))
    bw = gray.point(lambda p: 255 if p > OCR_BIN_THRESHOLD else 0)
    return bw


def _run_tesseract(img: Image.Image, config: str = "") -> "tuple[str, float]":
    """Roda o Tesseract em uma imagem já pronta e devolve (texto, confiança_média)."""
    data = pytesseract.image_to_data(
        img,
        lang=OCR_LANG,
        config=config,
        output_type=pytesseract.Output.DICT,
    )
    words, confs = [], []
    for word, conf in zip(data["text"], data["conf"]):
        word = word.strip()
        if not word:
            continue
        words.append(word)
        try:
            c = float(conf)
            if c >= 0:
                confs.append(c)
        except (TypeError, ValueError):
            pass
    text = " ".join(words)
    avg_conf = sum(confs) / len(confs) if confs else 0.0
    return text, avg_conf


def _ocr_image(img: Image.Image, two_pass: bool = True) -> "tuple[str, float]":
    tess_cmd = _tesseract_cmd()
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd

    text, avg_conf = _run_tesseract(img, config="")

    if two_pass:
        proc_img = _preprocess_for_ocr(img)
        text2, conf2 = _run_tesseract(proc_img, config=f"--psm {OCR_PSM}")
        text = f"{text}\n{text2}"
        avg_conf = max(avg_conf, conf2)

    return text, avg_conf


def _extract_via_ocr_pdf(
    pdf_path: str,
    first_page: int | None = None,
    last_page: int | None = None,
    two_pass: bool = True,
    dpi: int | None = None,
) -> "tuple[str, float, int]":
    # Limita o intervalo de páginas ANTES de renderizar.
    if first_page is None and last_page is None:
        try:
            info = pdfinfo_from_path(pdf_path, poppler_path=_poppler_path())
            total_paginas = info.get("Pages")
        except Exception:
            total_paginas = None
        if total_paginas:
            first_page = 1
            last_page = min(OCR_MAX_PAGES, total_paginas)

    images = convert_from_path(
        pdf_path,
        dpi=dpi or OCR_DPI,
        poppler_path=_poppler_path(),
        first_page=first_page,
        last_page=last_page,
    )

    if first_page is None and last_page is None and len(images) > OCR_MAX_PAGES:
        images = images[:OCR_MAX_PAGES]

    texts, confs = [], []

    for img in images:
        t, c = _ocr_image(img, two_pass=two_pass)

        texts.append(t)
        confs.append(c)

        if re.search(
            r"documento\s+assinado\s+digitalmente",
            t,
            re.I,
        ):
            print("Assinatura gov.br encontrada. OCR interrompido.")
            break

    full_text = "\n".join(texts).strip()
    avg_conf = sum(confs) / len(confs) if confs else 0.0

    return full_text, avg_conf, len(images)


def extract_ocr_forced(
    filepath: str,
    first_page: int | None = None,
    last_page: int | None = None,
    two_pass: bool = True,
    dpi: int | None = None,
) -> Optional[str]:
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".pdf":
            text, _conf, _n_pages = _extract_via_ocr_pdf(
                filepath,
                first_page=first_page,
                last_page=last_page,
                two_pass=two_pass,
                dpi=dpi,
            )
            return text
        elif ext in {".jpg", ".jpeg", ".png"}:
            text, _conf = _ocr_image(Image.open(filepath), two_pass=two_pass)
            return text
    except Exception:
        return None
    return None


def extract_ocr_forced_skip_edges(
    filepath: str,
    skip_first: int = 1,
    skip_last: int = 2,
    two_pass: bool = True,
    dpi: int | None = None,
) -> Optional[str]:
    """Como extract_ocr_forced, mas pulando as primeiras `skip_first` e
    últimas `skip_last` páginas. A assinatura gov.br (quando o sócio não
    assina via D4Sign) nunca aparece nesses trechos, então pular evita OCR
    à toa, principalmente quando a assinatura não é encontrada em lugar
    nenhum e o código hoje varre o PDF inteiro sem achar nada.

    dpi: opcional; quem chama pode pedir uma resolução mais baixa que a
    padrão."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext != ".pdf":
        return extract_ocr_forced(filepath, two_pass=two_pass)
    try:
        info = pdfinfo_from_path(filepath, poppler_path=_poppler_path())
        n_pages = info.get("Pages")
    except Exception:
        n_pages = None

    first_page = last_page = None
    if n_pages and n_pages > (skip_first + skip_last):
        first_page = skip_first + 1
        last_page = n_pages - skip_last

    return extract_ocr_forced(
        filepath, first_page=first_page, last_page=last_page, two_pass=two_pass, dpi=dpi
    )


def extract_ocr_progressivo(
    filepath: str, deve_parar, max_pages: int | None = None, two_pass: bool = False
) -> "tuple[str, float, int]":
    """OCR página a página (em vez de renderizar até OCR_MAX_PAGES
    páginas de uma vez), parando assim que `deve_parar(texto_acumulado_até_
    aqui)` devolver True.

    Usada para Contrato Social/Ata/Alteração: CNPJ e nome empresarial
    normalmente já estão na 1ª página, então não vale a pena OCR nas
    páginas seguintes. Se `deve_parar` nunca for satisfeito, processa até
    max_pages páginas (mesmo teto de OCR_MAX_PAGES usado no resto do app).
    """
    max_pages = max_pages or OCR_MAX_PAGES
    try:
        info = pdfinfo_from_path(filepath, poppler_path=_poppler_path())
        total_paginas = info.get("Pages") or 1
    except Exception:
        total_paginas = 1

    limite = min(max_pages, total_paginas)
    textos, confs = [], []
    for pagina in range(1, limite + 1):
        imagens = convert_from_path(
            filepath,
            dpi=OCR_DPI,
            poppler_path=_poppler_path(),
            first_page=pagina,
            last_page=pagina,
        )
        if not imagens:
            break
        t, c = _ocr_image(imagens[0], two_pass=two_pass)
        textos.append(t)
        confs.append(c)
        if deve_parar("\n".join(textos).strip()):
            break

    full_text = "\n".join(textos).strip()
    avg_conf = sum(confs) / len(confs) if confs else 0.0
    return full_text, avg_conf, len(textos)


def extract_progressivo(
    filepath: str, deve_parar, max_pages: int | None = None, two_pass: bool = False
) -> ExtractionResult:
    """Como extract(), mas quando o documento não tem texto nativo e precisa
    cair pra OCR, usa extract_ocr_progressivo() (página a página, com parada
    antecipada) em vez de sempre processar até OCR_MAX_PAGES páginas de uma
    vez. Documentos com texto nativo suficiente não passam por OCR, igual a
    extract()."""
    filename = os.path.basename(filepath)
    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == ".pdf":
            native = _extract_native_text(filepath)
            if len(native) >= MIN_NATIVE_TEXT_CHARS:
                with pdfplumber.open(filepath) as pdf:
                    n_pages = len(pdf.pages)
                return ExtractionResult(
                    filename=filename,
                    text=native,
                    method="native_text",
                    pages=n_pages,
                )
            text, conf, n_pages = extract_ocr_progressivo(
                filepath,
                deve_parar,
                max_pages=max_pages,
                two_pass=two_pass,
            )
            return ExtractionResult(
                filename=filename,
                text=text,
                method="ocr",
                ocr_confidence=round(conf, 1),
                pages=n_pages,
            )

        elif ext in {".jpg", ".jpeg", ".png", ".tif", ".tiff"}:
            img = Image.open(filepath)
            text, conf = _ocr_image(img, two_pass=two_pass)
            return ExtractionResult(
                filename=filename,
                text=text,
                method="ocr",
                ocr_confidence=round(conf, 1),
                pages=1,
            )

        else:
            return ExtractionResult(
                filename=filename,
                method="failed",
                error=f"Extensão não suportada: {ext}",
            )

    except Exception as exc:
        return ExtractionResult(filename=filename, method="failed", error=str(exc))


def extract(filepath: str, two_pass: bool = False) -> ExtractionResult:
    """Ponto de entrada único: recebe o caminho de um arquivo (PDF ou
    imagem) e devolve o texto extraído pela camada mais 'barata' que
    funcionar."""
    filename = os.path.basename(filepath)
    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == ".pdf":
            native = _extract_native_text(filepath)
            if len(native) >= MIN_NATIVE_TEXT_CHARS:
                with pdfplumber.open(filepath) as pdf:
                    n_pages = len(pdf.pages)
                return ExtractionResult(
                    filename=filename,
                    text=native,
                    method="native_text",
                    pages=n_pages,
                )
            # sem texto nativo suficiente -> cai para OCR
            text, conf, n_pages = _extract_via_ocr_pdf(filepath, two_pass=two_pass)
            return ExtractionResult(
                filename=filename,
                text=text,
                method="ocr",
                ocr_confidence=round(conf, 1),
                pages=n_pages,
            )

        elif ext in {".jpg", ".jpeg", ".png"}:
            img = Image.open(filepath)
            text, conf = _ocr_image(img, two_pass=two_pass)
            return ExtractionResult(
                filename=filename,
                text=text,
                method="ocr",
                ocr_confidence=round(conf, 1),
                pages=1,
            )

        else:
            return ExtractionResult(
                filename=filename,
                method="failed",
                error=f"Extensão não suportada: {ext}",
            )

    except Exception as exc:
        return ExtractionResult(filename=filename, method="failed", error=str(exc))
