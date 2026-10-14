"""Design tokens (design_system/theme.yaml): colors, fonts, type scale, spacing, logo, footer."""
from __future__ import annotations

import copy
import hashlib
import json
import re

from .config import _read_yaml, load_config
from .paths import design_dir, resolve_path

_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")
EMU_IN = 914400

DEFAULT_THEME = {
    "colors": {
        "primary": "1E4E8C",
        "secondary": "E07A1F",
        "chart": ["1E4E8C", "E07A1F", "5B8DB8", "7A8794", "A9C4DE", "C9A227"],
        "neutral": {"900": "1F2933", "700": "3E4C59", "500": "616E7C", "300": "CBD2D9", "100": "E4E7EB", "50": "F5F7FA"},
        "status": {"good": "2E7D32", "warn": "E0A100", "bad": "C62828"},
    },
    "fonts": {"latin": "Segoe UI", "ea": "Microsoft JhengHei"},
    "type_scale": {
        "cover": 40, "section": 32, "title": 26, "subtitle": 18, "headline": 20, "body": [18, 16, 14],
        "kpi_value": 40, "kpi_label": 13, "table": 12, "chart": 11, "note": 11, "caption": 10,
        "steps": [9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 44, 48, 54],
    },
    "spacing": {"margin_x": 0.6, "margin_top": 0.4, "gap": 0.2},
    "logo": {"path": "", "height": 0.35},
    "footer": {"text": "", "page_number": True},
}

STATUS_WORDS = {
    "good": ("good", "green", "ok", "on track", "on_track", u"綠", u"綠燈", u"正常", u"達標", u"完成", u"順利"),
    "warn": ("warn", "warning", "amber", "yellow", "risk", "at risk", u"黃", u"黃燈", u"注意", u"風險", u"觀察"),
    "bad": ("bad", "red", "off track", "off_track", "late", u"紅", u"紅燈", u"延遲", u"落後", u"異常", u"停滯"),
}


def _merge(base, over):
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        elif v is not None and v != "":
            out[k] = v
    return out


def _hex(v, default):
    s = str(v or "").strip().lstrip("#")
    return s.upper() if _HEX.match(s) else default


def status_key(value):
    """Map free text (綠/注意/red/…) to good|warn|bad, or None."""
    s = str(value or "").strip().lower()
    if not s:
        return None
    for key, words in STATUS_WORDS.items():
        if s in words:
            return key
    for key in ("bad", "warn", "good"):
        if any(w in s for w in STATUS_WORDS[key] if not w.isascii()):
            return key
    return None


def rgb_tuple(hex6):
    return tuple(int(hex6[i:i + 2], 16) for i in (0, 2, 4))


def _lum(hex6):
    def ch(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(c) for c in rgb_tuple(hex6))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(fg, bg) -> float:
    a, b = _lum(fg), _lum(bg)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def mix(hex_a, hex_b, t: float) -> str:
    """t=0 → a, t=1 → b."""
    a, b = rgb_tuple(hex_a), rgb_tuple(hex_b)
    return "".join("%02X" % int(round(x + (y - x) * t)) for x, y in zip(a, b))


class Theme:
    def __init__(self, data=None, cfg=None):
        cfg = cfg or {}
        d = _merge(DEFAULT_THEME, data or {})
        if not (data or {}).get("fonts", {}).get("ea") and cfg.get("builtin_font_ea"):
            d["fonts"]["ea"] = cfg["builtin_font_ea"]
        c, dc = d["colors"], DEFAULT_THEME["colors"]
        c["primary"] = _hex(c.get("primary"), dc["primary"])
        c["secondary"] = _hex(c.get("secondary"), dc["secondary"])
        chart = [_hex(x, "") for x in (c.get("chart") or [])]
        c["chart"] = [x for x in chart if x] or [c["primary"], c["secondary"]] + dc["chart"][2:]
        c["neutral"] = {str(k): _hex(v, dc["neutral"].get(str(k), "7B8794")) for k, v in c["neutral"].items()}
        c["status"] = {k: _hex(c["status"].get(k), v) for k, v in dc["status"].items()}
        # alert text colour: colors.alert > status.bad when theme.yaml exists > legacy config.json emphasis_colors.alert
        legacy = c["status"]["bad"] if data else _hex((cfg.get("emphasis_colors") or {}).get("alert"), c["status"]["bad"])
        c["alert"] = _hex(c.get("alert"), legacy)
        ts = d["type_scale"]
        body = ts.get("body") or [18, 16, 14]
        if not isinstance(body, list):
            body = [float(body), float(body) - 2, float(body) - 4]
        ts["body"] = [float(x) for x in (list(body) + [body[-1]] * 3)[:3]]
        ts["steps"] = sorted(set(float(x) for x in ts.get("steps") or DEFAULT_THEME["type_scale"]["steps"]))
        self.data = d

    # -- colors ---------------------------------------------------------
    @property
    def primary(self):
        return self.data["colors"]["primary"]

    @property
    def secondary(self):
        return self.data["colors"]["secondary"]

    @property
    def chart_colors(self):
        return list(self.data["colors"]["chart"])

    @property
    def alert(self):
        return self.data["colors"]["alert"]

    def neutral(self, level) -> str:
        n = self.data["colors"]["neutral"]
        return n.get(str(level)) or DEFAULT_THEME["colors"]["neutral"].get(str(level), "7B8794")

    def status(self, key) -> str:
        return self.data["colors"]["status"].get(key or "", self.neutral(500))

    def status_text(self, key) -> str:
        """Status colour darkened until it is readable as text on white / light cards (amber usually needs it)."""
        c = self.status(key)
        t = 0.0
        while contrast(c, self.neutral(50)) < 4.5 and t < 0.6:
            t += 0.05
            c = mix(self.status(key), "000000", t)
        return c

    # -- type -----------------------------------------------------------
    @property
    def font_latin(self):
        return self.data["fonts"].get("latin") or "Segoe UI"

    @property
    def font_ea(self):
        return self.data["fonts"].get("ea") or "Microsoft JhengHei"

    def size(self, role, level=0) -> float:
        v = self.data["type_scale"].get(role)
        if isinstance(v, list):
            return float(v[min(level, len(v) - 1)])
        return float(v) if v else 14.0

    @property
    def steps(self):
        return self.data["type_scale"]["steps"]

    def snap(self, size: float, floor: float = 0.0) -> float:
        """Largest step <= size (never below floor) so every font size in a deck comes from one scale."""
        cands = [s for s in self.steps if s <= size + 1e-6]
        v = cands[-1] if cands else self.steps[0]
        return max(v, floor) if floor else v

    # -- spacing / chrome -----------------------------------------------
    def inch(self, key) -> int:
        return int(float(self.data["spacing"].get(key, 0)) * EMU_IN)

    @property
    def footer_text(self):
        return str(self.data["footer"].get("text") or "")

    @property
    def page_number(self):
        return self.data["footer"].get("page_number", True) is not False

    def logo_path(self):
        p = str(self.data["logo"].get("path") or "").strip()
        if not p:
            return None
        path = resolve_path(p, base=design_dir().parent)
        return path if path.exists() else None

    @property
    def logo_height(self) -> int:
        return int(float(self.data["logo"].get("height") or 0.35) * EMU_IN)

    def fingerprint(self) -> str:
        logo = self.logo_path()
        extra = [str(logo), int(logo.stat().st_mtime)] if logo else []
        raw = json.dumps([self.data, extra], sort_keys=True, ensure_ascii=False)
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def load_theme(cfg=None) -> Theme:
    cfg = cfg if cfg is not None else load_config()
    p = design_dir() / "theme.yaml"
    data = _read_yaml(p) if p.exists() else None
    return Theme(data, cfg)
