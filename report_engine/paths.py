from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ENGINE_DIR = Path(__file__).resolve().parent


def design_dir() -> Path:
    env = os.environ.get("REPORT_DESIGN_DIR")
    return Path(env).expanduser().resolve() if env else ROOT / "design_system"


def output_dir() -> Path:
    env = os.environ.get("REPORT_OUTPUT_DIR")
    return Path(env).expanduser() if env else Path.cwd() / "output"


def resolve_path(p, base=None) -> Path:
    """Resolve a user-supplied path: absolute, else relative to base, cwd, then the repo root."""
    path = Path(str(p).strip().strip('"')).expanduser()
    if path.is_absolute():
        return path
    candidates = []
    if base:
        candidates.append(Path(base) / path)
    candidates.append(Path.cwd() / path)
    candidates.append(ROOT / path)
    for c in candidates:
        if c.exists():
            return c.resolve()
    return candidates[0].resolve()
