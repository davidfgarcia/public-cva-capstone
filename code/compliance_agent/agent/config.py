import os
from pathlib import Path
import yaml

AGENT_DIR = Path(__file__).resolve().parent.parent


def load_config(path=None, root=None):
    path = Path(path or os.environ.get("COMPLIANCE_CONFIG") or AGENT_DIR / "config.yaml")
    with open(path) as f:
        cfg = yaml.safe_load(f)
    root = root or os.environ.get("COMPLIANCE_ROOT") or cfg.get("root", "..")
    root = Path(root)
    if not root.is_absolute():
        root = (path.parent / root).resolve()
    cfg["root"] = root
    cfg["config_path"] = path
    cfg["projects_path"] = projects_folder(root, cfg.get("projects_dir", "Projects"))
    cfg.setdefault("projects", {})
    from .admin_settings import apply_config        # the Settings page's choices (#151), when its file is there
    apply_config(cfg)
    return cfg


def projects_folder(root, name="Projects"):
    """root/<name>; while the folder still has its old name (e.g. 'projects' before it was renamed 'Projects'),
    the folder whose name differs only in case."""
    root = Path(root)
    exact = root / name
    try:
        names = os.listdir(root)
    except OSError:
        return exact
    if name in names:
        return exact
    for n in names:
        if n.lower() == name.lower() and (root / n).is_dir():
            return root / n
    return exact


def project_cfg(cfg, project_name):
    return (cfg.get("projects") or {}).get(project_name, {}) or {}
