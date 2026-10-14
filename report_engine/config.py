"""Design-system files: config, audiences, style guide/rules, saved blueprints, examples."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Tuple

import yaml

from .paths import design_dir

DEFAULT_CONFIG = {
    "template": "",
    "default_audience": "manager",
    "builtin_font_ea": "Microsoft JhengHei",
    "emphasis_colors": {"alert": "C00000"},
    "min_font_pt": 12,
    "output_dir": "",
}


def read_json(path, default):
    try:
        with open(str(path), "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def write_json(path, data) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(str(tmp), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def _read_yaml(path) -> dict:
    with open(str(path), "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    cfg.update(read_json(design_dir() / "config.json", {}))
    return cfg


def save_config(cfg: dict) -> None:
    write_json(design_dir() / "config.json", cfg)


def list_audiences() -> Dict[str, dict]:
    out = {}
    for p in sorted((design_dir() / "audiences").glob("*.y*ml")):
        data = _read_yaml(p)
        data.setdefault("name", p.stem)
        out[p.stem] = data
    return out


def get_audience(name=None) -> Tuple[str, dict]:
    auds = list_audiences()
    if not name:
        name = load_config().get("default_audience") or "manager"
    name = str(name).strip()
    if name in auds:
        return name, auds[name]
    for key, prof in auds.items():
        aliases = [str(a).lower() for a in (prof.get("aliases") or [])]
        if name == prof.get("name") or name.lower() in aliases:
            return key, prof
    raise ValueError("找不到受眾「%s」，可用：%s" % (name, "、".join(auds) or "（audiences 資料夾是空的）"))


def style_guide_text() -> str:
    p = design_dir() / "style_guide.md"
    return p.read_text(encoding="utf-8") if p.exists() else ""


def style_rules() -> dict:
    p = design_dir() / "style_rules.yaml"
    return _read_yaml(p) if p.exists() else {}


def load_blueprints() -> dict:
    return read_json(design_dir() / "blueprints.json", {})


def save_blueprints(data: dict) -> None:
    write_json(design_dir() / "blueprints.json", data)


def examples() -> List[Tuple[Path, dict]]:
    out = []
    for p in sorted((design_dir() / "examples").glob("*.json"), key=lambda q: (-q.stat().st_mtime, q.name)):
        try:
            out.append((p, read_json(p, {})))
        except ValueError:
            continue
    return out


INTERNAL_KEYS = ("_build", "_history", "_output")


def save_example(spec_path, name) -> Path:
    """Store a finished, approved spec as a few-shot example for get_design_context."""
    import re

    spec = read_json(spec_path, None)
    if not isinstance(spec, dict) or not spec.get("slides"):
        raise ValueError("讀不到 spec：%s" % spec_path)
    clean = {k: v for k, v in spec.items() if k not in INTERNAL_KEYS}
    root = design_dir().parent
    srcs = []
    for s in clean.get("sources") or []:
        try:
            srcs.append(str(Path(s).resolve().relative_to(root)).replace("\\", "/"))
        except ValueError:
            srcs.append(s)
    clean["sources"] = srcs
    aud = str(clean.get("audience") or "manager")
    safe = re.sub(r'[\\/:*?"<>|\s]+', "_", str(name)).strip("_") or "example"
    dest = design_dir() / "examples" / ("%s_%s.json" % (aud, safe))
    write_json(dest, clean)
    return dest
