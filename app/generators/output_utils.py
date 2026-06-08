"""Output file naming and directory utilities."""

from __future__ import annotations

import os
from pathlib import Path


def generate_output_filename(project: str, month: str, year: int) -> str:
    """Return filename following ``<Project>_<Month>_<Year>.pptx`` convention.

    Example::

        >>> generate_output_filename("DSD_Mukkamala", "February", 2026)
        'DSD_Mukkamala_February_2026.pptx'
    """
    return f"{project}_{month}_{year}.pptx"


def ensure_output_directory(path: str) -> None:
    """Create the parent directory for *path* if it does not exist."""
    parent = Path(path).parent
    if parent != Path(".") and parent != Path(""):
        parent.mkdir(parents=True, exist_ok=True)


def resolve_output_path(output_arg: str, project: str, month: str, year: int) -> str:
    """Resolve the final output file path.

    * If *output_arg* is an existing directory **or** ends with a path
      separator, generate the filename and join it to the directory.
    * Otherwise treat *output_arg* as a full file path.

    In both cases the parent directory is created when missing and any
    existing file at the resolved path will be overwritten on save.
    """
    p = Path(output_arg)

    if p.is_dir() or output_arg.endswith(os.sep) or output_arg.endswith("/"):
        resolved = str(p / generate_output_filename(project, month, year))
    else:
        resolved = str(p)

    ensure_output_directory(resolved)
    return resolved
