"""Render normalized slides onto the template: placeholders first, grid regions for sketched layouts."""
from __future__ import annotations

import math
import re
import unicodedata

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .template import open_presentation, remove_all_slides
from .template_inspector import TEXT_KINDS, TITLE_TYPES, inspect_presentation

EMU_PT = 12700
EMU_IN = 914400
GUTTER = int(EMU_IN * 0.08)
RICH_RE = re.compile(r"(\*\*.+?\*\*|!!.+?!!)")
_NUMERIC_CELL = re.compile(r"^[+\-]?[\d,]+(\.\d+)?\s*%?$")
_BU_SUCCESSORS = (qn("a:tabLst"), qn("a:defRPr"), qn("a:extLst"))

CHART_MAP = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}


def plain_text(s) -> str:
    return RICH_RE.sub(lambda m: m.group(0)[2:-2], s or "")


def em_width(s: str) -> float:
    w = 0.0
    for ch in s:
        if unicodedata.east_asian_width(ch) in ("W", "F"):
            w += 1.0
        elif ch == " ":
            w += 0.28
        elif ch.isupper() or ch.isdigit():
            w += 0.6
        else:
            w += 0.5
    return w


def _text_height(items, width_pt: float) -> float:
    """items: (text, size_pt, indent_pt). Rough line-wrap estimate, deliberately a bit pessimistic."""
    total = 0.0
    for text, size, indent in items:
        avail = max(width_pt - indent, size * 3)
        lines = 0
        for seg in (text or " ").split("\n"):
            lines += max(1, int(math.ceil(em_width(seg) * size / avail)))
        total += lines * size * 1.2 + size * 0.3
    return total


def _remove(shape) -> None:
    el = shape._element
    el.getparent().remove(el)


def _lst_size(el, lvl):
    if el is None:
        return None
    d = el.find(".//%s/%s/%s" % (qn("a:lstStyle"), qn("a:lvl%dpPr" % lvl), qn("a:defRPr")))
    if d is not None and d.get("sz"):
        return int(d.get("sz")) / 100.0
    return None


def _master_sizes(master):
    tx = master._element.find(qn("p:txStyles"))

    def get(style, lvl, default):
        if tx is None:
            return default
        st = tx.find(qn(style))
        if st is None:
            return default
        d = st.find("%s/%s" % (qn("a:lvl%dpPr" % lvl), qn("a:defRPr")))
        return int(d.get("sz")) / 100.0 if d is not None and d.get("sz") else default

    title = get("p:titleStyle", 1, 36.0)
    body = [get("p:bodyStyle", i, d) for i, d in ((1, 24.0), (2, 20.0), (3, 18.0))]
    return title, body


def _insert_bu(pPr, el) -> None:
    for i, child in enumerate(pPr):
        if child.tag in _BU_SUCCESSORS:
            pPr.insert(i, el)
            return
    pPr.append(el)


def _clear_bu(pPr) -> None:
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buBlip", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)


def _no_bullet(para) -> None:
    pPr = para._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    _clear_bu(pPr)
    _insert_bu(pPr, etree.Element(qn("a:buNone")))


def _set_bullet(para, level: int, size: float) -> None:
    pPr = para._p.get_or_add_pPr()
    ind = int(size * EMU_PT * 1.1)
    pPr.set("marL", str(ind * (level + 1)))
    pPr.set("indent", str(-ind))
    _clear_bu(pPr)
    bu_font = etree.Element(qn("a:buFont"))
    bu_font.set("typeface", "Arial")
    bu_char = etree.Element(qn("a:buChar"))
    bu_char.set("char", u"•" if level == 0 else u"–")
    _insert_bu(pPr, bu_font)
    _insert_bu(pPr, bu_char)


def _plain_shape(shp) -> None:
    style = shp._element.find(qn("p:style"))
    if style is not None:
        shp._element.remove(style)
    shp.fill.solid()
    shp.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    shp.line.fill.background()


def _muted(font) -> None:
    font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    font.color.brightness = 0.4


class DeckRenderer:
    def __init__(self, template_path, cfg):
        self.cfg = cfg
        self.prs = open_presentation(template_path)
        remove_all_slides(self.prs)
        self.info = inspect_presentation(self.prs)
        self.by_key = {L["key"]: L for L in self.info["layouts"]}
        self.sw, self.sh = int(self.prs.slide_width), int(self.prs.slide_height)
        self.min_pt = float(cfg.get("min_font_pt", 12))
        alert = str((cfg.get("emphasis_colors") or {}).get("alert", "C00000")).lstrip("#")
        self.alert_rgb = RGBColor.from_string(alert)
        self.warnings = []
        self._where = ""
        master = self.prs.slide_masters[0]
        self.title_pt, self.body_pts = _master_sizes(master)
        self._calc_area(master)

    # -- geometry -------------------------------------------------------
    def _calc_area(self, master) -> None:
        mx, mt = int(self.sw * 0.05), int(self.sh * 0.06)
        bottom = self.sh - mt
        for ph in master.placeholders:
            t = ph.placeholder_format.type
            if ph.left is None or ph.top is None:
                continue
            if t in TITLE_TYPES:
                if 0 < ph.left < self.sw * 0.2:
                    mx = int(ph.left)
                if 0 < ph.top < self.sh * 0.2:
                    mt = int(ph.top)
            if t in (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE) and ph.top > self.sh * 0.75:
                bottom = min(bottom, int(ph.top) - GUTTER)
        self.area = (mx, mt, self.sw - 2 * mx, bottom - mt)

    def grid_rect(self, g):
        ax, ay, aw, ah = self.area
        x, y, w, h = g
        return (int(ax + x / 12.0 * aw) + GUTTER // 2,
                int(ay + y / 8.0 * ah) + GUTTER // 2,
                max(int(w / 12.0 * aw) - GUTTER, EMU_IN // 4),
                max(int(h / 8.0 * ah) - GUTTER, EMU_IN // 4))

    def _title_bottom_grid(self) -> float:
        L = self.by_key[self.info["with_title"]]
        if not L.get("title_bottom"):
            return 0.0
        ay, ah = self.area[1], self.area[3]
        return min(3.0, max(0.0, (L["title_bottom"] - ay) / float(ah) * 8 + 0.15))

    # -- entry ----------------------------------------------------------
    def render_all(self, nslides):
        out = []
        for ns in nslides:
            start = len(self.prs.slides)
            self._where = ns.get("where", "")
            try:
                if ns["mode"] == "custom":
                    self._custom(ns, ns["regions"])
                else:
                    self._layout_slide(ns)
            except Exception as e:  # keep the rest of the deck
                self.warnings.append("%s：繪製失敗（%s: %s）" % (self._where, type(e).__name__, e))
            out.append(list(range(start + 1, len(self.prs.slides) + 1)))
        return out

    def save(self, path) -> None:
        self.prs.save(str(path))

    # -- layout (placeholder) slides -----------------------------------
    @staticmethod
    def _headline_only(blocks) -> bool:
        return (len(blocks) == 1 and blocks[0]["type"] == "text"
                and all(p.get("headline") for p in blocks[0]["paras"]))

    def _choose(self, ns):
        blocks = ns["blocks"]
        key = ns.get("layout")
        if key:
            L = self.by_key.get(key)
            if L is None:
                for k, v in self.by_key.items():
                    if k.lower() == str(key).lower():
                        L = v
            if L is not None:
                return L
            self.warnings.append("%s：範本沒有版型「%s」，已自動挑選" % (self._where, key))
        cover = self.info.get("cover")
        if cover and (not blocks or self._headline_only(blocks)) and ns["idx"] == 1 and not ns.get("label"):
            return self.by_key[cover]
        k = self.info["by_count"].get(max(1, len(blocks)))
        return self.by_key[k] if k else None

    def _layout_slide(self, ns) -> None:
        L = self._choose(ns)
        blocks = list(ns["blocks"])
        cover_fit = L is not None and L.get("subtitle_idx") is not None and self._headline_only(blocks)
        if L is None or (len(blocks) > len(L["content_idx"]) and not cover_fit):
            if L is not None:
                self.warnings.append("%s：內容區塊（%d）多於版型「%s」的內容區（%d），已自動改用自訂排版"
                                     % (self._where, len(blocks), L["key"], len(L["content_idx"])))
            else:
                self.warnings.append("%s：範本沒有剛好 %d 個內容區的版型，已自動改用自訂排版" % (self._where, len(blocks)))
            self._custom(ns, self._auto_regions(blocks, bool(ns.get("title"))))
            return
        slide = self.prs.slides.add_slide(L["obj"])
        phs = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
        title_ph = phs.get(L["title_idx"]) if L["title_idx"] is not None else None
        if title_ph is not None:
            if ns.get("title"):
                self._set_title(title_ph, ns["title"], L["obj"])
            else:
                _remove(title_ph)
        elif ns.get("title"):
            self.warnings.append("%s：版型「%s」沒有標題區，標題未放入" % (self._where, L["key"]))
        sub_ph = phs.get(L["subtitle_idx"]) if L.get("subtitle_idx") is not None else None
        if sub_ph is not None:
            if cover_fit:
                self._fill_text_ph(sub_ph, blocks[0]["paras"], L["obj"], allow_split=False, plain_headline=True)
                blocks = []
            else:
                _remove(sub_ph)
        content_phs = [phs[i] for i in L["content_idx"] if i in phs]
        kinds = [c["kind"] for c in L["content"]]
        leftover = None
        for i, (ph, block) in enumerate(zip(content_phs, blocks)):
            if block["type"] == "text" and kinds[i] in TEXT_KINDS:
                leftover = self._fill_text_ph(ph, block["paras"], L["obj"], allow_split=True)
            else:
                rect = (int(ph.left), int(ph.top), int(ph.width), int(ph.height))
                _remove(ph)
                self._draw(slide, rect, block)
        for ph in content_phs[len(blocks):]:
            _remove(ph)
        self._notes(slide, ns.get("notes"))
        guard = 0
        while leftover and guard < 6:
            guard += 1
            leftover = self._continuation(ns, L, leftover)

    def _continuation(self, ns, L, paras):
        slide = self.prs.slides.add_slide(L["obj"])
        phs = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
        title_ph = phs.get(L["title_idx"]) if L["title_idx"] is not None else None
        if title_ph is not None:
            self._set_title(title_ph, (ns.get("title") or "") + u"（續）", L["obj"])
        if L.get("subtitle_idx") in phs:
            _remove(phs[L["subtitle_idx"]])
        content_phs = [phs[i] for i in L["content_idx"] if i in phs]
        if not content_phs:
            return None
        left = self._fill_text_ph(content_phs[0], paras, L["obj"], allow_split=True)
        for ph in content_phs[1:]:
            _remove(ph)
        return left

    def _ph_sizes(self, ph, layout, default):
        idx = ph.placeholder_format.idx
        sizes = list(default)
        lph = None
        for p in layout.placeholders:
            if p.placeholder_format.idx == idx:
                lph = p
                break
        for el in ((lph._element if lph is not None else None), ph._element):
            for lvl in (1, 2, 3):
                sz = _lst_size(el, lvl)
                if sz:
                    sizes[lvl - 1] = sz
        for i in (1, 2):
            sizes[i] = min(sizes[i], sizes[i - 1])
        return sizes

    def _fit(self, height_fn, h_pt, base0, min_pt=None):
        min_s = min(1.0, (min_pt or self.min_pt) / float(base0))
        s = 1.0
        while s >= min_s - 1e-9:
            if height_fn(s) <= h_pt:
                return s
            s -= 0.05
        return None

    def _set_title(self, ph, text, layout) -> None:
        base = self._ph_sizes(ph, layout, [self.title_pt] * 3)[0]
        w_pt = ph.width / float(EMU_PT) - 14.4
        h_pt = ph.height / float(EMU_PT) - 7.2
        plain = plain_text(text)

        def height(s):
            size = base * s
            lines = max(1, int(math.ceil(em_width(plain) * size / max(w_pt, size))))
            return lines * size * 1.1

        scale = self._fit(height, h_pt, base, min_pt=18)
        if scale is None:
            scale = min(1.0, 18.0 / base)
            self.warnings.append("%s：標題太長，建議精簡到 30 字內" % self._where)
        tf = ph.text_frame
        tf.clear()
        self._runs(tf.paragraphs[0], text, base * scale if scale < 0.999 else None)

    def _fill_text_ph(self, ph, paras, layout, allow_split, plain_headline=False):
        base = self._ph_sizes(ph, layout, self.body_pts)
        w_pt = ph.width / float(EMU_PT) - 14.4
        h_pt = (ph.height / float(EMU_PT) - 7.2) * 0.95

        def items(ps, s):
            return [(plain_text(p["text"]), base[min(p["level"], 2)] * s,
                     (p["level"] + 1) * 20.0 if p.get("bullet", True) and not p.get("headline") else 0.0)
                    for p in ps]

        scale = self._fit(lambda s: _text_height(items(paras, s), w_pt), h_pt, base[0])
        leftover = None
        if scale is None:
            min_s = min(1.0, self.min_pt / base[0])
            if allow_split and len(paras) > 1:
                k = len(paras)
                while k > 1 and _text_height(items(paras[:k], min_s), w_pt) > h_pt:
                    k -= 1
                leftover, paras = paras[k:], paras[:k]
                scale = self._fit(lambda s: _text_height(items(paras, s), w_pt), h_pt, base[0]) or min_s
            else:
                scale = min_s
                self.warnings.append("%s：文字過多，已縮到最小字級，仍可能超出版面" % self._where)
        explicit = scale < 0.999
        tf = ph.text_frame
        tf.clear()
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            lvl = min(p["level"], 2)
            if lvl:
                para.level = lvl
            if p.get("headline") or not p.get("bullet", True):
                _no_bullet(para)
            size = base[lvl] * scale if explicit else None
            self._runs(para, p["text"], size, bold=bool(p.get("headline")) and not plain_headline, emph=p.get("emph"))
        return leftover

    # -- custom (grid) slides ------------------------------------------
    def _auto_regions(self, blocks, has_title):
        top = self._title_bottom_grid() if has_title else 0.0
        H = 8.0 - top
        regs = []
        kp = [b for b in blocks if b["type"] == "kpis"]
        rest = [b for b in blocks if b["type"] != "kpis"]
        if kp and rest:
            band = min(2.0, H * 0.3)
            w = 12.0 / len(kp)
            for i, b in enumerate(kp):
                regs.append(((i * w, top, w, band), b))
            top += band
            H -= band
        elif kp:
            rest = kp
        n = len(rest)
        if n == 1:
            regs.append(((0, top, 12, H), rest[0]))
        elif n == 2:
            a = 5 if rest[0]["type"] == "text" else 6
            regs += [((0, top, a, H), rest[0]), ((a, top, 12 - a, H), rest[1])]
        elif n == 3:
            regs += [((0, top, 4, H), rest[0]), ((4, top, 8, H / 2), rest[1]), ((4, top + H / 2, 8, H / 2), rest[2])]
        elif n > 3:
            rows = int(math.ceil(n / 2.0))
            for i, b in enumerate(rest):
                regs.append((((i % 2) * 6, top + (i // 2) * H / rows, 6, H / rows), b))
        return [{"rect": r, "block": b} for r, b in regs]

    def _custom(self, ns, regions) -> None:
        has_title = bool(ns.get("title"))
        L = self.by_key[self.info["with_title"] if has_title else self.info["blank"]]
        slide = self.prs.slides.add_slide(L["obj"])
        title_bottom = self.area[1]
        for ph in list(slide.placeholders):
            if has_title and ph.placeholder_format.type in TITLE_TYPES and title_bottom == self.area[1]:
                self._set_title(ph, ns["title"], L["obj"])
                title_bottom = int(ph.top + ph.height)
            else:
                _remove(ph)
        for reg in regions:
            x, y, w, h = self.grid_rect(reg["rect"])
            if has_title and y < title_bottom + GUTTER // 2:
                h = max(h - (title_bottom + GUTTER - y), EMU_IN // 2)
                y = title_bottom + GUTTER
            self._draw(slide, (x, y, w, h), reg["block"])
        self._notes(slide, ns.get("notes"))

    # -- blocks ---------------------------------------------------------
    def _draw(self, slide, rect, block) -> None:
        t = block["type"]
        if t == "text":
            self._draw_text(slide, rect, block["paras"], (18.0, 16.0, 14.0))
        elif t == "headline":
            self._draw_text(slide, rect, [{"text": block["text"], "level": 0, "bullet": False, "headline": True}],
                            (24.0, 24.0, 24.0), anchor=MSO_ANCHOR.MIDDLE, min_pt=14)
        elif t == "note":
            self._draw_text(slide, rect, [{"text": block["text"], "level": 0, "bullet": False}],
                            (11.0, 11.0, 11.0), min_pt=9, muted=True)
        elif t == "kpis":
            self._draw_kpis(slide, rect, block["items"])
        elif t == "chart":
            self._draw_chart(slide, rect, block)
        elif t == "table":
            self._draw_table(slide, rect, block)
        elif t == "image":
            self._draw_image(slide, rect, block)

    def _textbox(self, slide, rect, anchor=MSO_ANCHOR.TOP):
        x, y, w, h = rect
        tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = Emu(91440)
        tf.margin_top = tf.margin_bottom = Emu(45720)
        tf.vertical_anchor = anchor
        return tb, tf

    def _draw_text(self, slide, rect, paras, base, anchor=MSO_ANCHOR.TOP, min_pt=None, muted=False) -> None:
        w_pt = rect[2] / float(EMU_PT) - 14.4
        h_pt = (rect[3] / float(EMU_PT) - 7.2) * 0.95

        def height(s):
            return _text_height([(plain_text(p["text"]), base[min(p["level"], 2)] * s,
                                  (p["level"] + 1) * base[min(p["level"], 2)] * s * 1.1 if p.get("bullet", True) else 0.0)
                                 for p in paras], w_pt)

        scale = self._fit(height, h_pt, base[0], min_pt)
        if scale is None:
            scale = min(1.0, (min_pt or self.min_pt) / base[0])
            self.warnings.append("%s：區塊文字過多，已縮到最小字級，仍可能超出框線" % self._where)
        _tb, tf = self._textbox(slide, rect, anchor)
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            lvl = min(p["level"], 2)
            size = base[lvl] * scale
            if i:
                para.space_before = Pt(round(size * 0.3, 1))
            if p.get("bullet", True) and not p.get("headline"):
                _set_bullet(para, lvl, size)
            self._runs(para, p["text"], size, bold=bool(p.get("headline")), emph=p.get("emph"),
                       color=_muted if muted else None)

    def _runs(self, para, text, size=None, bold=False, emph=None, color=None) -> None:
        for part in RICH_RE.split(text or ""):
            if not part:
                continue
            e = emph
            if part.startswith("**") and part.endswith("**") and len(part) >= 4:
                part, e = part[2:-2], (emph or "accent")
            elif part.startswith("!!") and part.endswith("!!") and len(part) >= 4:
                part, e = part[2:-2], "alert"
            r = para.add_run()
            r.text = part.replace("\n", " ")
            f = r.font
            if size:
                f.size = Pt(round(size * 2) / 2.0)
            if bold or e:
                f.bold = True
            if e in ("accent", "accent1"):
                f.color.theme_color = MSO_THEME_COLOR.ACCENT_1
            elif e == "accent2":
                f.color.theme_color = MSO_THEME_COLOR.ACCENT_2
            elif e in ("alert", "red"):
                f.color.rgb = self.alert_rgb
            elif color is not None:
                color(f)

    def _draw_kpis(self, slide, rect, items) -> None:
        n = len(items)
        if not n:
            return
        x, y, w, h = rect
        best = None
        for cols in range(1, n + 1):
            rows = int(math.ceil(n / float(cols)))
            score = abs(math.log(((w / float(cols)) / (h / float(rows))) / 1.8))
            if best is None or score < best[0]:
                best = (score, cols, rows)
        _s, cols, rows = best
        cw, ch = w // cols, h // rows
        for i, it in enumerate(items):
            cx = x + (i % cols) * cw
            cy = y + (i // cols) * ch
            self._kpi_card(slide, (cx + GUTTER // 2, cy + GUTTER // 2, cw - GUTTER, ch - GUTTER), it)

    def _kpi_card(self, slide, rect, it) -> None:
        x, y, w, h = rect
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y),
                                     Emu(min(w, int(EMU_IN * 0.6))), Emu(int(EMU_IN * 0.05)))
        _plain_shape(bar)
        w_pt = w / float(EMU_PT) - 14.4
        h_pt = h / float(EMU_PT)
        val = plain_text(it["value"])
        vsize = max(16.0, min(54.0, h_pt * 0.42, w_pt * 0.95 / max(em_width(val), 1.0)))
        lsize = max(11.0, min(16.0, vsize * 0.36))
        _tb, tf = self._textbox(slide, (x, y + int(EMU_IN * 0.08), w, h - int(EMU_IN * 0.08)), MSO_ANCHOR.MIDDLE)
        self._runs(tf.paragraphs[0], it["value"], vsize, bold=True, emph=it.get("emph") or "accent")
        if it.get("label"):
            self._runs(tf.add_paragraph(), it["label"], lsize)
        if it.get("delta"):
            self._runs(tf.add_paragraph(), it["delta"], lsize, color=_muted)

    def _draw_chart(self, slide, rect, b) -> None:
        x, y, w, h = rect
        ctype = b["chart_type"]
        series = b["series"]
        pie = ctype in ("pie", "doughnut")
        if pie and len(series) > 1:
            self.warnings.append("%s：圓餅圖只能有一組數列，已只用「%s」" % (self._where, series[0]["name"]))
            series = series[:1]
        pct = all(s.get("percent") for s in series)
        data = CategoryChartData(number_format="0%" if pct else "General")
        data.categories = [str(c) for c in b["categories"]]
        for s in series:
            data.add_series(str(s["name"] or " "), [None if v is None else float(v) for v in s["values"]])
        chart = slide.shapes.add_chart(CHART_MAP[ctype], Emu(x), Emu(y), Emu(w), Emu(h), data).chart
        chart.font.size = Pt(12)
        multi = len(series) > 1 or pie
        chart.has_legend = multi
        if multi:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
        if b.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = plain_text(b["title"])
            chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        else:
            chart.has_title = False
        if pct and not pie:
            chart.value_axis.tick_labels.number_format = "0%"
            chart.value_axis.tick_labels.number_format_is_linked = False
        if b.get("labels") or pie:
            plot = chart.plots[0]
            plot.has_data_labels = True
            plot.data_labels.font.size = Pt(11)
            if pct:
                plot.data_labels.number_format = "0%"
                plot.data_labels.number_format_is_linked = False
        if ctype in ("column", "bar", "stacked", "stacked_bar"):
            chart.plots[0].gap_width = 80

    def _cell(self, cell, text, size, bold=False, emph=None, right=False) -> None:
        cell.margin_left = cell.margin_right = Emu(54864)
        cell.margin_top = cell.margin_bottom = Emu(27432)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        if right:
            p.alignment = PP_ALIGN.RIGHT
        self._runs(p, text, size, bold=bold, emph=emph)

    def _draw_table(self, slide, rect, b) -> None:
        cols, rows = b["columns"], b["rows"]
        if not cols:
            return
        x, y, w, h = rect
        h_pt = h / float(EMU_PT)
        font = 12.0

        def cap(f):
            return max(1, int(h_pt // (f * 2.0)) - 1)

        while font > 9 and len(rows) > cap(font):
            font -= 1
        if len(rows) > cap(font):
            self.warnings.append("%s：表格 %d 列放不下，只顯示前 %d 列（可用 limit 篩選或拆頁）"
                                 % (self._where, len(rows), cap(font)))
            rows = rows[:cap(font)]
        nrows = len(rows) + 1
        row_h = int(font * 2.0 * EMU_PT)
        tbl = slide.shapes.add_table(nrows, len(cols), Emu(x), Emu(y), Emu(w), Emu(min(h, row_h * nrows))).table
        weights = []
        for j, c in enumerate(cols):
            texts = [plain_text(str(c))] + [plain_text(str(r[j])) if j < len(r) and r[j] is not None else "" for r in rows]
            weights.append(min(24.0, max(3.0, max(em_width(t) for t in texts))))
        total, used = sum(weights), 0
        for j, wt in enumerate(weights):
            cw = int(w * wt / total) if j < len(weights) - 1 else w - used
            tbl.columns[j].width = Emu(cw)
            used += cw
        for i in range(nrows):
            tbl.rows[i].height = Emu(row_h)
        for j, c in enumerate(cols):
            self._cell(tbl.cell(0, j), str(c), font, bold=True)
        highlight = set(b.get("highlight") or ())
        for i, r in enumerate(rows, 1):
            for j in range(len(cols)):
                v = r[j] if j < len(r) and r[j] is not None else ""
                v = str(v)
                self._cell(tbl.cell(i, j), v, font, emph="alert" if (i - 1) in highlight else None,
                           right=bool(_NUMERIC_CELL.match(v)))

    def _draw_image(self, slide, rect, b) -> None:
        from PIL import Image

        x, y, w, h = rect
        with Image.open(str(b["path"])) as im:
            iw, ih = im.size
        s = min(w / float(iw), h / float(ih))
        pw, ph = int(iw * s), int(ih * s)
        slide.shapes.add_picture(str(b["path"]), Emu(x + (w - pw) // 2), Emu(y + (h - ph) // 2), Emu(pw), Emu(ph))

    def _notes(self, slide, notes) -> None:
        if not notes:
            return
        tf = slide.notes_slide.notes_text_frame
        if tf is not None:
            tf.text = notes
