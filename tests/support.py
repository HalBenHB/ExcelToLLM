from __future__ import annotations

import shutil
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4


@contextmanager
def workspace_temp_dir(prefix: str):
    tmp_root = Path(__file__).resolve().parent / "_tmp"
    tmp_root.mkdir(parents=True, exist_ok=True)
    workdir = tmp_root / f"{prefix}_{uuid4().hex[:8]}"
    workdir.mkdir(parents=True, exist_ok=True)
    try:
        yield workdir
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
