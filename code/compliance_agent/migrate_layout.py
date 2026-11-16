"""One-off: move every project in the shared Projects folder from an older folder layout to the current one
(Uploads, Payroll, Tracker, Files). Runs are never allowed to do this: a project still in an old layout stops the
run for that project, with a message pointing here.

    python3 migrate_layout.py "/path/to/compliance/Projects"            # dry run: changes nothing
    python3 migrate_layout.py "/path/to/compliance/Projects" --apply    # does it

Dry run: copies the Projects folder to a temporary folder, migrates the copy, checks that every file is still
there unchanged, and prints where each file would go. Nothing in the real folder is touched.

--apply: the same, on the real folder.
* Files are only moved or renamed, never changed or deleted. The one exception: .agent_state.json, whose saved
  paths are updated to the new folders.
* The old folders that are left empty are moved, not deleted, to
  <compliance root>/_sync_backup/<time>__layout_migration/old folders/<project>/. That folder also gets a list of
  every file before and after, and the report. _sync_backup is never copied to AWS. Delete it by hand when you're
  satisfied.
* Before it finishes, it checks that every file from before is still present with the same content: same checksum,
  in Projects or in the old-folders backup.
* Run it with the Projects folder on Google Drive for desktop. Moving files there keeps their Drive IDs, so the
  tokenizer moves the AWS copies the same way.

Safe to run again: projects already in the current layout are left alone."""
from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import sys
import tempfile
import time
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from agent.layout import Paths, _free_name, _merge_dir, ensure_layout, migrate_project, old_layout_dirs   # noqa: E402

SKIP_NAMES = {".DS_Store"}
CHANGED_BY_DESIGN = {".agent_state.json"}      # its saved paths are updated to the new folders


def manifest(*roots: Path):
    """[(path relative to its root, size, md5)] for every file under the roots."""
    out = []
    for root in roots:
        if not root.is_dir():
            continue
        for f in sorted(root.rglob("*")):
            if f.is_file() and f.name not in SKIP_NAMES:
                h = hashlib.md5()
                with open(f, "rb") as fh:
                    for chunk in iter(lambda: fh.read(1 << 20), b""):
                        h.update(chunk)
                out.append((f"{root.name}/{f.relative_to(root).as_posix()}", f.stat().st_size, h.hexdigest()))
    return out


def projects_in(projects_dir: Path):
    return sorted(p for p in projects_dir.iterdir() if p.is_dir() and not p.name.startswith((".", "_")))


def needs_migration(p: Path) -> bool:
    return bool(old_layout_dirs(p, with_files_only=False))


def migrate(projects_dir: Path, hold: Path, note=print):
    """Migrate every project under projects_dir. Files the layout code leaves in an old folder (e.g. in a subfolder
    of Upload/) go to Files/Unused/<old folder>/ inside the project; the emptied old folders then go to
    hold/'old folders'/<project>/. The file list before is written first, and each project's result as it is done
    (note). Returns (before, after) manifests: after = Projects + whatever ended up in hold."""
    before = manifest(projects_dir)
    write_tsv(hold / "before.tsv", before)
    for p in projects_in(projects_dir):
        if not needs_migration(p):
            ensure_layout(p)
            note(f"{p.name}: already in the current layout")
            continue
        moved, _ = migrate_project(p)
        left = old_layout_dirs(p, with_files_only=False)
        extra = []
        for name in left:
            src = p / name
            stray = [f for f in src.rglob("*") if f.is_file() and f.name not in SKIP_NAMES]
            if stray:              # not placed by the layout rules (subfolders, loose files): kept in the project
                n = _merge_dir(src, Paths(p).unused / name)
                extra.append(f"{n} file(s) left in '{name}' (subfolders or loose files) moved to Files/Unused/{name}/")
            if src.exists():
                dest_dir = hold / "old folders" / p.name
                dest_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(src), str(_free_name(dest_dir, name)))
        ensure_layout(p)
        note(f"{p.name}: {'; '.join(moved) or 'folders created'}"
             + "".join(f"; NOTE {x}" for x in extra)
             + (f"; emptied old folders set aside: {', '.join(left)}" if left else ""))
    after = manifest(projects_dir, hold / "old folders")
    return before, after


def check(before, after):
    """Problems: files from before missing after (by content), and unexpected new files."""
    key = lambda e: (e[1], e[2])
    b = Counter(key(e) for e in before if Path(e[0]).name not in CHANGED_BY_DESIGN)
    a = Counter(key(e) for e in after if Path(e[0]).name not in CHANGED_BY_DESIGN)
    problems = []
    held = [e[0] for e in after if e[0].startswith("old folders/")]
    if held:
        problems.append(f"{len(held)} file(s) ended up in the old-folders backup instead of the project: "
                        + "; ".join(held[:20]))
    missing = b - a
    if missing:
        names = [e[0] for e in before if key(e) in missing]
        problems.append(f"{sum(missing.values())} file(s) not found unchanged after the move: " + "; ".join(names[:20]))
    extra = a - b
    if extra:
        names = [e[0] for e in after if key(e) in extra]
        problems.append(f"{sum(extra.values())} unexpected new file(s): " + "; ".join(names[:20]))
    nb = sum(1 for e in before if Path(e[0]).name in CHANGED_BY_DESIGN)
    na = sum(1 for e in after if Path(e[0]).name in CHANGED_BY_DESIGN)
    if nb != na:
        problems.append(f".agent_state.json files: {nb} before, {na} after")
    return problems


def moves(before, after):
    """[(old path, new path)] for files whose path changed, matched by content (and name when there are copies)."""
    pool = defaultdict(list)
    for rel, size, md5 in after:
        pool[(size, md5)].append(rel)
    out = []
    for rel, size, md5 in before:
        cands = pool.get((size, md5)) or []
        if not cands:
            continue
        stem = Path(rel).stem
        pick = next((c for c in cands if Path(c).name == Path(rel).name), None) \
            or next((c for c in cands if Path(c).stem.startswith(stem + " (")), cands[0])
        cands.remove(pick)
        if pick != rel:
            out.append((rel, pick))
    return out


def tree_summary(projects_dir: Path):
    out = []
    for p in projects_in(projects_dir):
        parts = []
        for d in sorted(x for x in p.iterdir() if x.is_dir() and not x.name.startswith(".")):
            n = sum(1 for f in d.rglob("*") if f.is_file() and f.name not in SKIP_NAMES)
            parts.append(f"{d.name} ({n})")
        out.append(f"  {p.name}: {' | '.join(parts)}")
    return out


def write_tsv(path: Path, rows):
    path.write_text("path\tbytes\tmd5\n" + "".join(f"{r}\t{s}\t{m}\n" for r, s, m in rows))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("projects", type=Path, help="the shared Projects folder")
    ap.add_argument("--apply", action="store_true", help="migrate the real folder (default: a dry run on a copy)")
    ap.add_argument("--backup", type=Path, help="where the old folders and the report go (default: _sync_backup "
                                                "next to the Projects folder)")
    a = ap.parse_args(argv)
    projects_dir = a.projects.expanduser().resolve()
    if not projects_dir.is_dir():
        sys.exit(f"Not a folder: {projects_dir}")
    todo = [p.name for p in projects_in(projects_dir) if needs_migration(p)]
    print(f"Projects folder: {projects_dir}")
    print(f"{len(projects_in(projects_dir))} project(s); in an old layout: {', '.join(todo) or 'none'}")
    if not todo:
        print("Nothing to do.")
        return 0
    stamp = time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
    if a.apply:
        base = a.backup.expanduser().resolve() if a.backup else projects_dir.parent / "_sync_backup"
        probe = base
        while not probe.exists():
            probe = probe.parent
        if os.stat(probe).st_dev != os.stat(projects_dir).st_dev:
            # a move to another drive is a copy + delete; deleting may be refused, which would leave copies
            sys.exit(f"Refusing to run: the backup folder {base} is not on the same drive as {projects_dir}. "
                     "Use --backup with a folder on that drive (e.g. next to Projects).")
        hold = base / f"{stamp}__layout_migration"
        hold.mkdir(parents=True, exist_ok=True)
        work = projects_dir
    else:
        tmp = Path(tempfile.mkdtemp(prefix="layout_dry_run_"))
        work = tmp / projects_dir.name
        print(f"Dry run: copying to {work} ...")
        shutil.copytree(projects_dir, work, symlinks=True)
        hold = tmp / "_sync_backup" / f"{stamp}__layout_migration"
        hold.mkdir(parents=True)
    report_file = hold / "report.txt"
    head = ["DRY RUN (nothing in the real folder was changed)" if not a.apply else "APPLIED",
            f"Projects folder: {projects_dir}", ""]
    report_file.write_text("\n".join(head) + "\n")

    def note(line):                       # each project's result is saved as soon as it is done
        print(line)
        with open(report_file, "a") as fh:
            fh.write(line + "\n")

    print("\n".join(head))
    if a.apply:
        before, after = migrate(work, hold, note)
    else:                                 # like the real run from Claude's shell: deleting isn't allowed
        def refuse(*_a, **_k):
            raise PermissionError("deleting is not allowed (dry run behaves like the real run)")
        real = Path.rmdir, Path.unlink
        Path.rmdir = Path.unlink = refuse
        try:
            before, after = migrate(work, hold, note)
        finally:
            Path.rmdir, Path.unlink = real
    problems = check(before, after)
    mv = moves(before, after)
    tail = ["", f"{len(before)} file(s) before, {len(after)} after (Projects + old-folders backup); "
                f"{len(mv)} moved or renamed", "", "Folders now (files in each):"] + tree_summary(work)
    tail += ["", "CHECK: every file is still there, unchanged" if not problems else "CHECK FAILED:"]
    tail += [f"  {x}" for x in problems]
    tail += ["", "Moves:"] + [f"  {o}  ->  {n}" for o, n in mv]
    write_tsv(hold / "after.tsv", after)
    with open(report_file, "a") as fh:
        fh.write("\n".join(tail) + "\n")
    print("\n".join(tail))
    print(f"\nReport and file lists: {hold}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
