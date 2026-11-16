"""Saving a workbook so that a run killed mid-save (out of memory, say) never leaves a broken file behind: it is
written to '<name>.tmp' (never synced: deploy/rclone-filter.txt) and then renamed over the real one in one step."""
from __future__ import annotations

import os
from pathlib import Path


def save_workbook(wb, path) -> None:
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            tmp.unlink()
