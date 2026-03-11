"""
App-level helpers (folder scan, etc.). No changes to src/.
"""
from pathlib import Path
from typing import List, Union

from app.config import SUPPORTED_EXTENSIONS


def get_supported_files_from_folder(
    folder_path: Union[str, Path],
    include_subfolders: bool = False,
) -> List[Path]:
    """
    Return a flat list of supported file paths under folder_path.
    Excludes Excel lock files (~$*).
    """
    folder = Path(folder_path)
    if not folder.is_dir():
        return []

    files: List[Path] = []
    if include_subfolders:
        for ext in SUPPORTED_EXTENSIONS:
            files.extend(folder.rglob(f"*{ext}"))
    else:
        for ext in SUPPORTED_EXTENSIONS:
            files.extend(folder.glob(f"*{ext}"))

    return [p for p in sorted(files) if not p.name.startswith("~$")]
