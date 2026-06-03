"""
PACE Payroll Validator — File Retriever
========================================
Locates payroll files on disk, copies them into the local dataset
folder, and renames them using the DocName from the database.
"""

import shutil
from pathlib import Path
from typing import Dict, List

from config import DATASET_DIR
from logger_setup import setup_logger

logger = setup_logger()


def retrieve_files(
    woid: str,
    file_info_list: List[Dict[str, str]],
    dataset_dir: Path = DATASET_DIR,
) -> List[Path]:
    """
    Copy and rename payroll files into ``dataset/<WOID>/``.

    Parameters
    ----------
    woid : str
        The Work Order ID.
    file_info_list : List[Dict[str, str]]
        Each dict must have ``DocName`` and ``RecosSpec`` keys.
    dataset_dir : Path
        Root dataset directory (default from config).

    Returns
    -------
    List[Path]
        Paths to the successfully copied files.
    """
    woid_dir = dataset_dir / woid
    woid_dir.mkdir(parents=True, exist_ok=True)

    copied_files: List[Path] = []

    for info in file_info_list:
        doc_name = info["DocName"]
        recos_spec = info["RecosSpec"]

        source_file = _locate_source_file(recos_spec, woid)
        if source_file is None:
            continue

        dest_path = woid_dir / doc_name
        try:
            shutil.copy2(str(source_file), str(dest_path))
            logger.info(
                "WOID %s - Copied %s -> %s", woid, source_file.name, dest_path.name
            )
            copied_files.append(dest_path)
        except OSError as exc:
            logger.error(
                "WOID %s — Failed to copy %s: %s", woid, source_file, exc
            )

    if not copied_files:
        logger.warning("WOID %s — No files were copied.", woid)

    return copied_files


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _locate_source_file(recos_spec: str, woid: str):
    """
    Resolve a RecosSpec path to an actual file on disk.

    RecosSpec is a full path *without* an extension, e.g.
    ``D:\\PACE\\DATA\\12345``.  The physical file may or may not
    have an extension.

    Strategy:
        1. Try the path exactly as given (no extension).
        2. Search the parent folder for any file whose stem matches.
    """
    spec_path = Path(recos_spec)
    folder = spec_path.parent
    file_stem = spec_path.name          # e.g. "12345"

    # Check folder exists
    if not folder.is_dir():
        logger.error(
            "WOID %s — Folder does not exist: %s", woid, folder
        )
        return None

    # 1) Exact match (file without extension)
    if spec_path.is_file():
        return spec_path

    # 2) Search for any file with matching stem
    matches = [f for f in folder.iterdir() if f.is_file() and f.stem == file_stem]

    if len(matches) == 1:
        return matches[0]
    elif len(matches) > 1:
        # If multiple matches, prefer the one without extension first,
        # otherwise take the first alphabetically.
        no_ext = [f for f in matches if f.suffix == ""]
        chosen = no_ext[0] if no_ext else sorted(matches)[0]
        logger.warning(
            "WOID %s — Multiple files match stem '%s'; using %s",
            woid, file_stem, chosen.name,
        )
        return chosen
    else:
        logger.error(
            "WOID %s — File not found for RecosSpec: %s", woid, recos_spec
        )
        return None
