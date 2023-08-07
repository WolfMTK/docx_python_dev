from __future__ import annotations
from pathlib import Path


def default_docx_path() -> Path:
    return Path(__file__) / 'templates' / 'default.docx'
