"""SVG → native, editable PowerPoint shapes (no rasterising, no extra packages).

Supported: svg/g (transform, inherited style), rect (rx), circle, ellipse, line, polyline, polygon,
path (M L H V C S Q T A Z, absolute + relative; curves are flattened), text/tspan, marker-end/-start arrows,
fill/stroke/stroke-width/stroke-dasharray/opacity, style="…", linearGradient/radialGradient (first stop colour),
and theme colour names (primary, secondary, alert, good/warn/bad, neutral-50…900, chart-1…6, var(--primary)).
"""
from __future__ import annotations

import math
import re

from lxml import etree

from .canvas import EMU_PT, Canvas, resolve_color, set_alt

SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False, remove_comments=True)
_NUM = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
_TOKEN = re.compile(r"([MmLlHhVvCcSsQqTtAaZz])|(%s)" % _NUM)
_TF = re.compile(r"(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)")
_URL = re.compile(r"url\(\s*#([^)\s]+)\s*\)")
INHERIT = ("fill", "stroke", "stroke-width", "stroke-dasharray", "fill-opacity", "stroke-opacity",
           "font-size", "font-weight", "text-anchor", "dominant-baseline")
SKIP = {"defs", "title", "desc", "metadata", "lineargradient", "radialgradient", "stop", "marker", "symbol"}
IGNORED = {"style", "use", "image", "pattern", "clippath", "mask", "filter", "foreignobject", "switch", "a"}
CURVE_STEPS = 16


def local(tag):
    return tag.rsplit("}", 1)[-1].lower() if isinstance(tag, str) else ""


def parse_svg(markup):
    """-> root element; raises ValueError with a readable message."""
    if not isinstance(markup, str) or "<svg" not in markup:
        raise ValueError(u"內容要是一段 <svg …>…</svg>")
    try:
        root = etree.fromstring(markup.strip().encode("utf-8"), SAFE_PARSER)
    except etree.XMLSyntaxError as e:
        raise ValueError(u"SVG 語法錯誤：%s" % e)
    if local(root.tag) != "svg":
        raise ValueError(u"最外層要是 <svg>")
    return root


def _nums(s):
    return [float(x) for x in re.findall(_NUM, s or "")]


def _num(v, default=0.0):
    if v is None:
        return default
    m = re.match(r"\s*(%s)" % _NUM, str(v))
    return float(m.group(1)) if m else default


# -- affine matrices (a b c d e f) ------------------------------------------
def _mul(m, n):
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f)


def _apply(m, x, y):
    a, b, c, d, e, f = m
    return a * x + c * y + e, b * x + d * y + f


def _transform(s):
    m = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for name, args in _TF.findall(s or ""):
        v = _nums(args)
        if name == "matrix" and len(v) == 6:
            t = tuple(v)
        elif name == "translate" and v:
            t = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif name == "scale" and v:
            t = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif name == "rotate" and v:
            r = math.radians(v[0])
            t = (math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0, 0)
            if len(v) == 3:
                t = _mul(_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        elif name == "skewX" and v:
            t = (1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        elif name == "skewY" and v:
            t = (1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
        else:
            continue
        m = _mul(m, t)
    return m


def _axis_aligned(m):
    return abs(m[1]) < 1e-9 and abs(m[2]) < 1e-9


def _scale(m):
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2])) or 1.0


# -- path data ----------------------------------------------------------------
def _arc(x1, y1, rx, ry, phi, large, sweep, x2, y2):
    """Endpoint arc → points (SVG implementation notes F.6)."""
    if rx == 0 or ry == 0:
        return [(x2, y2)]
    rx, ry = abs(rx), abs(ry)
    p = math.radians(phi)
    cp, sp = math.cos(p), math.sin(p)
    dx, dy = (x1 - x2) / 2.0, (y1 - y2) / 2.0
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = (x1p ** 2) / (rx ** 2) + (y1p ** 2) / (ry ** 2)
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
    den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
    k = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        k = -k
    cxp, cyp = k * rx * y1p / ry, -k * ry * x1p / rx
    cx, cy = cp * cxp - sp * cyp + (x1 + x2) / 2.0, sp * cxp + cp * cyp + (y1 + y2) / 2.0

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(4, int(abs(dt) / (math.pi / 12)))
    out = []
    for i in range(1, n + 1):
        t = t1 + dt * i / n
        out.append((cx + rx * math.cos(t) * cp - ry * math.sin(t) * sp, cy + rx * math.cos(t) * sp + ry * math.sin(t) * cp))
    return out


def _bez(p0, p1, p2, p3=None):
    out = []
    for i in range(1, CURVE_STEPS + 1):
        t = i / float(CURVE_STEPS)
        u = 1 - t
        if p3 is None:
            out.append((u * u * p0[0] + 2 * u * t * p1[0] + t * t * p2[0], u * u * p0[1] + 2 * u * t * p1[1] + t * t * p2[1]))
        else:
            out.append((u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
                        u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1]))
    return out


def path_points(d):
    """SVG path data → [(points, closed)] in user units."""
    toks = [(c, float(n) if n else None) for c, n in _TOKEN.findall(d or "")]
    subs, cur, start = [], [], (0.0, 0.0)
    x = y = 0.0
    cmd = None
    last_ctrl = None
    i = 0

    def take(k):
        nonlocal i
        vals = []
        while len(vals) < k:
            if i >= len(toks) or toks[i][0]:
                raise ValueError(u"path 參數不足")
            vals.append(toks[i][1])
            i += 1
        return vals

    while i < len(toks):
        c, _n = toks[i]
        if c:
            cmd = c
            i += 1
            if cmd in "Zz":
                if cur:
                    subs.append((cur, True))
                cur = []
                x, y = start
                last_ctrl = None
                continue
        elif cmd is None:
            raise ValueError(u"path 要以 M 開頭")
        rel = cmd.islower()
        C = cmd.upper()
        ox, oy = (x, y) if rel else (0.0, 0.0)
        if C == "M":
            px, py = take(2)
            if cur:
                subs.append((cur, False))
            x, y = ox + px, oy + py
            start = (x, y)
            cur = [(x, y)]
            cmd = "l" if rel else "L"
            last_ctrl = None
            continue
        if not cur:
            cur = [(x, y)]
        if C == "L":
            px, py = take(2)
            x, y = ox + px, oy + py
            cur.append((x, y))
            last_ctrl = None
        elif C == "H":
            (px,) = take(1)
            x = ox + px
            cur.append((x, y))
            last_ctrl = None
        elif C == "V":
            (py,) = take(1)
            y = oy + py
            cur.append((x, y))
            last_ctrl = None
        elif C in "CS":
            if C == "C":
                a, b, c2, d2, e, f = take(6)
                p1 = (ox + a, oy + b)
            else:
                c2, d2, e, f = take(4)
                p1 = (2 * x - last_ctrl[0], 2 * y - last_ctrl[1]) if last_ctrl else (x, y)
            p2, p3 = (ox + c2, oy + d2), (ox + e, oy + f)
            cur.extend(_bez((x, y), p1, p2, p3))
            last_ctrl = p2
            x, y = p3
        elif C in "QT":
            if C == "Q":
                a, b, e, f = take(4)
                p1 = (ox + a, oy + b)
            else:
                e, f = take(2)
                p1 = (2 * x - last_ctrl[0], 2 * y - last_ctrl[1]) if last_ctrl else (x, y)
            p2 = (ox + e, oy + f)
            cur.extend(_bez((x, y), p1, p2))
            last_ctrl = p1
            x, y = p2
        elif C == "A":
            rx, ry, phi, large, sweep, e, f = take(7)
            nx, ny = ox + e, oy + f
            cur.extend(_arc(x, y, rx, ry, phi, int(large), int(sweep), nx, ny))
            x, y = nx, ny
            last_ctrl = None
        else:
            i += 1
    if cur and len(cur) > 1:
        subs.append((cur, False))
    return subs


# -- converter --------------------------------------------------------------
class _Converter:
    def __init__(self, renderer, slide, rect, root):
        self.r = renderer
        self.theme = renderer.theme
        self.root = root
        self.canvas = Canvas(renderer, slide, "SVG")
        self.ignored = set()
        self.small_text = False
        self.gradients = {}
        for el in root.iter():
            if local(el.tag) in ("lineargradient", "radialgradient") and el.get("id"):
                stops = [s for s in el if local(s.tag) == "stop"]
                if stops:
                    st = self._style_attr(stops[0])
                    self.gradients[el.get("id")] = st.get("stop-color") or stops[0].get("stop-color")
        x, y, w, h = rect
        vb = _nums(root.get("viewBox") or root.get("viewbox"))
        if len(vb) == 4 and vb[2] > 0 and vb[3] > 0:
            vx, vy, vw, vh = vb
        else:
            vx, vy = 0.0, 0.0
            vw = _num(root.get("width"), 0) or w / float(EMU_PT)
            vh = _num(root.get("height"), 0) or h / float(EMU_PT)
        k = min(w / vw, h / vh)  # preserveAspectRatio xMidYMid meet
        ox = x + (w - vw * k) / 2.0 - vx * k
        oy = y + (h - vh * k) / 2.0 - vy * k
        self.base = (k, 0.0, 0.0, k, ox, oy)

    @staticmethod
    def _style_attr(el):
        out = {}
        for part in (el.get("style") or "").split(";"):
            if ":" in part:
                a, b = part.split(":", 1)
                out[a.strip().lower()] = b.strip()
        return out

    def _style(self, el, parent):
        st = {k: v for k, v in parent.items() if k in INHERIT}
        st["opacity"] = parent.get("opacity", 1.0)
        for k in INHERIT + ("marker-end", "marker-start"):
            if el.get(k) is not None:
                st[k] = el.get(k)
        css = self._style_attr(el)
        for k in INHERIT + ("marker-end", "marker-start"):
            if k in css:
                st[k] = css[k]
        op = css.get("opacity", el.get("opacity"))
        if op is not None:
            st["opacity"] = st["opacity"] * _num(op, 1.0)
        return st

    def _color(self, v):
        if v is None:
            return None
        m = _URL.match(str(v).strip())
        if m:
            v = self.gradients.get(m.group(1))
            if v is None:
                return None
        return resolve_color(v, self.theme)

    def _paint(self, st, m, default_fill="black"):
        fill = self._color(st.get("fill", default_fill))
        stroke = self._color(st.get("stroke"))
        width = _num(st.get("stroke-width"), 1.0) * _scale(m) / EMU_PT
        dash = bool(st.get("stroke-dasharray")) and str(st.get("stroke-dasharray")).strip() not in ("none", "0")
        op = float(st.get("opacity", 1.0))
        fa = op * _num(st.get("fill-opacity"), 1.0)
        sa = op * _num(st.get("stroke-opacity"), 1.0)
        return dict(fill=fill, stroke=stroke, width=width, dash=dash, alpha=fa, stroke_alpha=sa)

    def walk(self, el, m, parent):
        tag = local(el.tag)
        if tag in SKIP or not tag:
            return
        if tag in IGNORED:
            self.ignored.add(tag)
            return
        st = self._style(el, parent)
        if (el.get("display") or self._style_attr(el).get("display")) == "none":
            return
        if el.get("transform"):
            m = _mul(m, _transform(el.get("transform")))
        if tag in ("svg", "g"):
            if tag == "svg" and el is not self.root:
                m = _mul(m, (1, 0, 0, 1, _num(el.get("x")), _num(el.get("y"))))
            for child in el:
                self.walk(child, m, st)
            return
        getattr(self, "_" + tag, self._unknown)(el, m, st)

    def _unknown(self, el, m, st):
        self.ignored.add(local(el.tag))

    def _rect(self, el, m, st):
        x, y = _num(el.get("x")), _num(el.get("y"))
        w, h = _num(el.get("width")), _num(el.get("height"))
        if w <= 0 or h <= 0:
            return
        rx = _num(el.get("rx"), _num(el.get("ry"), 0))
        p = self._paint(st, m)
        if _axis_aligned(m):
            (x1, y1), (x2, y2) = _apply(m, x, y), _apply(m, x + w, y + h)
            kw = {k: p[k] for k in ("fill", "stroke", "width", "dash", "alpha")}
            self.canvas.rect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1), radius=rx * abs(m[0]), **kw)
        else:
            pts = [_apply(m, *q) for q in ((x, y), (x + w, y), (x + w, y + h), (x, y + h))]
            self.canvas.poly([(pts, True)], **p)

    def _ellipse_like(self, m, st, cx, cy, rx, ry):
        p = self._paint(st, m)
        if _axis_aligned(m):
            ccx, ccy = _apply(m, cx, cy)
            self.canvas.ellipse(ccx, ccy, abs(rx * m[0]), abs(ry * m[3]),
                                **{k: p[k] for k in ("fill", "stroke", "width", "dash", "alpha")})
        else:
            pts = [_apply(m, cx + rx * math.cos(t * math.pi / 24), cy + ry * math.sin(t * math.pi / 24)) for t in range(48)]
            self.canvas.poly([(pts, True)], **p)

    def _circle(self, el, m, st):
        r = _num(el.get("r"))
        if r > 0:
            self._ellipse_like(m, st, _num(el.get("cx")), _num(el.get("cy")), r, r)

    def _ellipse(self, el, m, st):
        rx, ry = _num(el.get("rx")), _num(el.get("ry"))
        if rx > 0 and ry > 0:
            self._ellipse_like(m, st, _num(el.get("cx")), _num(el.get("cy")), rx, ry)

    def _arrows(self, st):
        return dict(arrow_start=bool(st.get("marker-start")) and st.get("marker-start") != "none",
                    arrow_end=bool(st.get("marker-end")) and st.get("marker-end") != "none")

    def _line(self, el, m, st):
        p = self._paint(st, m, default_fill="none")
        x1, y1 = _apply(m, _num(el.get("x1")), _num(el.get("y1")))
        x2, y2 = _apply(m, _num(el.get("x2")), _num(el.get("y2")))
        self.canvas.line(x1, y1, x2, y2, p["stroke"], p["width"], p["dash"], alpha=p["stroke_alpha"], **self._arrows(st))

    def _points(self, el, m, st, closed):
        v = _nums(el.get("points"))
        pts = [_apply(m, v[i], v[i + 1]) for i in range(0, len(v) - 1, 2)]
        p = self._paint(st, m, default_fill="black" if closed else st.get("fill", "black"))
        if not closed and len(pts) == 2 and not p["fill"]:
            self.canvas.line(pts[0][0], pts[0][1], pts[1][0], pts[1][1], p["stroke"], p["width"], p["dash"],
                             alpha=p["stroke_alpha"], **self._arrows(st))
            return
        self.canvas.poly([(pts, closed)], **dict(p, **self._arrows(st)))

    def _polyline(self, el, m, st):
        self._points(el, m, st, False)

    def _polygon(self, el, m, st):
        self._points(el, m, st, True)

    def _path(self, el, m, st):
        subs = path_points(el.get("d"))
        subs = [([_apply(m, *q) for q in pts], closed) for pts, closed in subs]
        p = self._paint(st, m)
        if not p["fill"] and len(subs) == 1 and len(subs[0][0]) == 2:
            (a, b) = subs[0][0]
            self.canvas.line(a[0], a[1], b[0], b[1], p["stroke"], p["width"], p["dash"],
                             alpha=p["stroke_alpha"], **self._arrows(st))
            return
        self.canvas.poly(subs, **dict(p, **self._arrows(st)))

    def _text(self, el, m, st):
        runs = []
        base_x, base_y = _num(el.get("x")), _num(el.get("y"))
        if el.text and el.text.strip():
            runs.append((base_x, base_y, el.text.strip(), st))
        cy = base_y
        for ts in el:
            if local(ts.tag) != "tspan":
                continue
            tst = self._style(ts, st)
            tx = _num(ts.get("x"), base_x) + _num(ts.get("dx"), 0)
            cy = _num(ts.get("y"), cy) + _num(ts.get("dy"), 0)
            text = "".join(ts.itertext()).strip()
            if text:
                runs.append((tx, cy, text, tst))
            if ts.tail and ts.tail.strip():
                runs.append((tx, cy, ts.tail.strip(), st))
        for x, y, text, s in runs:
            px, py = _apply(m, x, y)
            size = _num(s.get("font-size"), 16) * _scale(m) / EMU_PT
            if size < 9:
                self.small_text = True
            anchor = {"middle": "middle", "end": "end"}.get(s.get("text-anchor"), "start")
            base = s.get("dominant-baseline") or ""
            valign = "middle" if base in ("middle", "central") else ("top" if base in ("hanging", "text-before-edge") else "baseline")
            weight = str(s.get("font-weight", ""))
            bold = weight in ("bold", "bolder") or (weight.isdigit() and int(weight) >= 600)
            color = self._color(s.get("fill", "black")) or self.theme.neutral(900)
            self.canvas.text(px, py, text, max(size, 9.0), color=color, bold=bold, align=anchor, valign=valign)


def draw_svg(renderer, slide, rect, markup):
    """Draw SVG markup inside rect (EMU). Returns warnings."""
    root = parse_svg(markup)
    conv = _Converter(renderer, slide, rect, root)
    conv.walk(root, conv.base, {"fill": "black", "opacity": 1.0})
    title = next((el.text for el in root.iter() if local(el.tag) == "title" and el.text), None)
    texts = [t.strip() for el in root.iter() if local(el.tag) in ("text", "tspan") for t in [el.text or ""] if t.strip()]
    set_alt(conv.canvas.group, u"示意圖：" + (title or u"、".join(texts[:12])))
    notes = []
    if conv.ignored:
        notes.append(u"SVG 不支援的元素已略過：%s" % u"、".join(sorted(conv.ignored)))
    if conv.small_text:
        notes.append(u"SVG 縮放後有文字小於 9pt，已放大到 9pt；請加大 font-size 或讓 viewBox 接近區塊實際大小")
    if conv.canvas.count == 0:
        notes.append(u"SVG 沒有畫出任何東西")
    return notes
