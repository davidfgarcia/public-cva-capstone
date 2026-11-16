"""Copy the registry (code/filekinds/cva_filekinds/) into both components.

    python code/filekinds/sync.py           copy
    python code/filekinds/sync.py --check   only report whether the copies match (exit 1 if not)

Each component's deploy builds only its own folder, so each carries a copy; a test in each fails if its copy differs
from the source (code/tokenizer/tests/test_filekinds.py, code/compliance_agent/tests/test_filekinds.py).
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "cva_filekinds"
CODE = HERE.parent
COPIES = (CODE / "tokenizer" / "pii_tokenizer" / "cva_filekinds", CODE / "compliance_agent" / "agent" / "cva_filekinds")
HEADER = "# GENERATED COPY of code/filekinds/cva_filekinds/{name}: edit the source, then run code/filekinds/sync.py\n"


def _files(root: Path) -> dict:
    """{relative path: bytes} of every file under ``root``, subfolders included (compiled caches left out)."""
    if not root.is_dir():
        return {}
    return {p.relative_to(root).as_posix(): p.read_bytes() for p in sorted(root.rglob("*"))
            if p.is_file() and "__pycache__" not in p.parts and p.suffix != ".pyc"
            and not any(part.startswith(".") for part in p.relative_to(root).parts)}   # .DS_Store and the like


def source_files() -> dict:
    return _files(SOURCE)


def expected(name: str, data: bytes) -> bytes:
    """A .py file gets the generated-copy header; any other file is copied as it is."""
    return HEADER.format(name=name).encode("utf-8") + data if name.endswith(".py") else data


def differences(copy: Path) -> list:
    """The files of ``copy`` that are missing, changed (byte for byte) or extra, against the source."""
    want = {n: expected(n, d) for n, d in source_files().items()}
    have = _files(copy)
    return sorted(n for n in set(want) | set(have) if want.get(n) != have.get(n))


def sync() -> None:
    files = source_files()
    for copy in COPIES:
        if copy.exists():
            shutil.rmtree(copy)
        copy.mkdir(parents=True)
        for name, data in files.items():
            (copy / name).parent.mkdir(parents=True, exist_ok=True)
            (copy / name).write_bytes(expected(name, data))


def main(argv) -> int:
    if "--check" in argv:
        bad = {str(c.relative_to(CODE)): differences(c) for c in COPIES}
        bad = {c: d for c, d in bad.items() if d}
        for c, d in bad.items():
            print(f"{c}: differs from the source in {', '.join(d)}")
        return 1 if bad else 0
    sync()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
