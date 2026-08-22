"""
Utilitários de descompactação/localização de pastas de OS dentro de um zip.
Compartilhados pelos dois fluxos (COBAN e Não COBAN); não conhecem nenhum
dos dois, só lidam com arquivos/pastas.
"""

import os
import re
import zipfile

from src.shared.constants import SUPPORTED_EXTENSIONS

OS_FOLDER_RE = re.compile(r"OS[_\s-]*(\d+)", re.I)


def extract_zip(zip_path: str, dest_dir: str) -> None:
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(dest_dir)


def find_os_folders(dest_dir: str) -> list[str]:
    """Um zip pode conter os documentos soltos numa pasta, ou várias
    subpastas (um lote de vários pacotes). Considera "pasta de documentos"
    qualquer diretório que contenha diretamente arquivos com extensão
    suportada."""
    candidate_dirs = []
    for root, _dirs, files in os.walk(dest_dir):
        if "__MACOSX" in root:
            continue
        has_supported_file = any(
            os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS for f in files
        )
        if has_supported_file:
            candidate_dirs.append(root)
    return candidate_dirs or [dest_dir]


def os_id_from_folder(folder_path: str, fallback: str) -> str:
    m = OS_FOLDER_RE.search(os.path.basename(folder_path))
    return m.group(1) if m else fallback
