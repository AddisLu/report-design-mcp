"""spec (JSON from the model) -> normalized slides -> .pptx, in one call. Also patch_deck."""
from __future__ import annotations

import datetime as _dt
import json
import re
from collections import OrderedDict

from lxml import etree
from pathlib import Path

from . import config as C
from .facts import BRACE_RE, RefError, Resolver, aggregate, format_value
from .paths import output_dir, resolve_path
from .renderer import DeckRenderer, plain_text
from .sources import load_sources
from .template import resolve_template_path
from .diagrams import to_date
from .svg import local, parse_svg
from .theme import status_key

REGION_TYPES = {
    "title": "title", u"標題": "title",
    "headline": "headline", u"結論": "headline", u"重點句": "headline",
    "bullets": "bullets", u"重點": "bullets", u"條列": "bullets",
    "text": "text", u"文字": "text", u"說明": "text",
    "kpi": "kpi", "kpis": "kpi", u"指標": "kpi", u"數字": "kpi",
    "chart": "chart", u"圖": "chart", u"圖表": "chart",
    "table": "table", u"表": "table", u"表格": "table",
    "image": "image", u"圖片": "image", u"照片": "image", u"截圖": "image", "screenshot": "image",
    "images": "images", "gallery": "images", u"圖組": "images", u"截圖組": "images", u"多張圖": "images",
    "note": "note", u"備註": "note", u"註": "note",
    "status": "status", u"狀態": "status", u"燈號": "status", u"紅綠燈": "status", u"紅黃綠燈": "status",
    "gantt": "gantt", u"甘特圖": "gantt", u"甘特": "gantt", u"排程": "gantt",
    "timeline": "timeline", u"時程表": "timeline", u"時程": "timeline", u"時間軸": "timeline", u"里程碑": "timeline",
    "process": "process", u"流程": "process", u"流程圖": "process", u"步驟": "process",
    "svg": "svg", "diagram": "svg", "drawing": "svg", u"示意圖": "svg", u"繪圖": "svg", u"圖解": "svg", u"架構圖": "svg",
}
DIAGRAM_KEYS = ("gantt", "timeline", "process", "svg", "images")
CHART_TYPES = {"column", "bar", "stacked", "stacked_bar", "line", "area", "pie", "doughnut"}
MAX_REPEAT = 40
_NUM_WARN_RE = re.compile(r"(?<![\w.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?\s*%|\d+\.\d+|\d{3,})(?!\w)")
_YEAR_RE = re.compile(r"^(19|20)\d{2}$")


def _deep_replace(obj, reps):
    if isinstance(obj, str):
        for a, b in reps.items():
            obj = obj.replace(a, b)
        return obj
    if isinstance(obj, list):
        return [_deep_replace(x, reps) for x in obj]
    if isinstance(obj, dict):
        return {k: _deep_replace(v, reps) for k, v in obj.items()}
    return obj


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _sort_rows(rows, spec_sort):
    desc = spec_sort.startswith("-")
    col = spec_sort.lstrip("-+").strip()
    have = [r for r in rows if r.get(col) is not None]
    none = [r for r in rows if r.get(col) is None]

    def key(r):
        v = r.get(col)
        return (0, v, "") if _is_num(v) else (1, 0, str(v))

    return sorted(have, key=key, reverse=desc) + none


def _safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", str(s)).strip(" ._")
    return (s or "report")[:80]


class Build:
    def __init__(self, spec, base_dir=None):
        self.spec = spec
        self.errors = []
        self.warnings = []
        self.cfg = C.load_config()
        self.rules = C.style_rules()
        self.replaced = 0
        try:
            self.aud_key, self.aud = C.get_audience(spec.get("audience"))
        except ValueError as e:
            self.errors.append(str(e))
            self.aud_key, self.aud = str(spec.get("audience") or "manager"), {}
        self.bundle = load_sources(spec.get("sources") or [], base_dir)
        self.errors.extend(self.bundle.errors)
        self.resolver = Resolver(self.bundle.tables, spec.get("facts") or {})
        self.blueprints = C.load_blueprints()

    # -- text -----------------------------------------------------------
    def _lint(self, raw, where):
        has_data = bool(self.bundle.tables)
        if not re.search(r"\d|\{", raw):
            for ph in self.rules.get("vague_phrases") or []:
                if ph in raw:
                    self.warnings.append(u"%s：「%s」偏空泛，建議改成具體數字或事實" % (where, ph))
                    break
        for ph in self.rules.get("banned_phrases") or []:
            if ph in raw:
                self.warnings.append(u"%s：用到禁用詞「%s」" % (where, ph))
        if has_data:
            stripped = BRACE_RE.sub("", raw)
            hits = [m.group(1) for m in _NUM_WARN_RE.finditer(stripped) if not _YEAR_RE.match(m.group(1))]
            if hits:
                self.warnings.append(u"%s：數字「%s」沒有引用資料，請確認來源或改用 {…}" % (where, "、".join(hits[:3])))

    def _terms(self, s):
        for a, b in (self.rules.get("replacements") or {}).items():
            if a in s:
                self.replaced += s.count(a)
                s = s.replace(a, b)
        return s

    def txt(self, value, where):
        if value is None:
            return None
        if not isinstance(value, str):
            value = format_value(value)
        self._lint(value, where)
        return self._terms(self.resolver.render(value, self.errors, where))

    def paras(self, items, where, level=0, out=None, bullet=True):
        out = [] if out is None else out
        if isinstance(items, str):
            items = [items]
        for it in items or []:
            if isinstance(it, list):
                self.paras(it, where, level + 1, out, bullet)
                continue
            if isinstance(it, dict):
                lvl = int(it.get("level", level))
                out.append({"text": self.txt(it.get("text", ""), where), "level": max(0, min(lvl, 2)),
                            "emph": it.get("emphasis"), "bullet": bullet})
                if it.get("children"):
                    self.paras(it["children"], where, lvl + 1, out, bullet)
            else:
                out.append({"text": self.txt(str(it), where), "level": min(level, 2), "emph": None, "bullet": bullet})
        return out

    def plain_paras(self, text, where):
        lines = [x for x in str(text).split("\n") if x.strip()] if isinstance(text, str) else (text or [])
        return self.paras(lines, where, bullet=False)

    # -- blocks ---------------------------------------------------------
    def kpis(self, items, where):
        if isinstance(items, dict):
            items = [items]
        res = []
        for it in items or []:
            if isinstance(it, str):
                it = {"value": it}
            val = it.get("value")
            if val is None and it.get("fact"):
                f = str(it["fact"]).strip()
                val = f if f.startswith("{") else "{%s}" % f
            good = str(it.get("good") or "up").lower()
            res.append({"value": self.txt(val, where) or u"—", "label": self.txt(it.get("label"), where) or "",
                        "delta": self.txt(it.get("delta"), where), "emph": it.get("emphasis"),
                        "good": "down" if good in ("down", u"低", u"下降", u"越低越好") else "up",
                        "status": status_key(self.txt(it.get("status"), where)) if it.get("status") else None})
        if len(res) > 4:
            self.warnings.append(u"%s：%d 個 KPI 偏多，建議 4 個以內，只留最重要的" % (where, len(res)))
        return {"type": "kpis", "items": res} if res else None

    def status(self, st, where):
        """Traffic-light list: [{name, status, note, owner}] or {from, name, status, note, owner} column names."""
        items = []
        if isinstance(st, dict) and st.get("from"):
            try:
                t, rows = self.resolver.rows_from(st["from"])
            except RefError as e:
                self.errors.append(u"%s：status.from → %s" % (where, e))
                return None
            cols = {k: st.get(k) for k in ("name", "status", "note", "owner") if st.get(k)}
            missing = [c for c in cols.values() if c not in t.columns]
            if "name" not in cols or "status" not in cols or missing:
                self.errors.append(u"%s：status 需要 name 與 status 欄位名稱，且欄位要存在（缺：%s；%s 欄位：%s）"
                                   % (where, u"、".join(missing) or "name/status", t.id, u"、".join(t.columns)))
                return None
            sort = str(st.get("sort") or "")
            if sort.lstrip("-+").strip() in ("status", cols["status"]):  # by severity: 延遲 → 注意 → 正常
                rank = {"bad": 0, "warn": 1, "good": 2}
                rows = sorted(rows, key=lambda r: rank.get(status_key(r.get(cols["status"])), 3),
                              reverse=sort.startswith("-"))
            elif sort:
                rows = _sort_rows(rows, sort)
            if st.get("limit"):
                rows = rows[:int(st["limit"])]
            for r in rows:
                raw = {k: format_value(r.get(c), t.kinds.get(c)) if r.get(c) is not None else "" for k, c in cols.items()}
                items.append(raw)
        else:
            for it in (st.get("items") if isinstance(st, dict) else st) or []:
                if isinstance(it, str):
                    it = {"name": it}
                items.append({k: self.txt(it.get(k), where) or "" for k in ("name", "status", "note", "owner")})
        out = []
        for it in items:
            key = status_key(it.get("status"))
            if it.get("status") and key is None:
                self.warnings.append(u"%s：狀態「%s」看不懂，請用 正常/注意/延遲（或 綠/黃/紅）" % (where, it["status"]))
            out.append({"name": self._terms(it.get("name") or ""), "status": key, "label": it.get("status") or "",
                        "note": self._terms(it.get("note") or ""), "owner": it.get("owner") or ""})
        if not out:
            self.errors.append(u"%s：status 沒有項目" % where)
            return None
        return {"type": "status", "items": out}

    def chart(self, c, where):
        if not isinstance(c, dict):
            self.errors.append(u"%s：chart 必須是物件" % where)
            return None
        ctype = str(c.get("type", "column")).lower()
        if ctype not in CHART_TYPES:
            self.warnings.append(u"%s：不認得圖表類型「%s」，改用 column" % (where, ctype))
            ctype = "column"
        cats, series = [], []
        if c.get("from"):
            try:
                t, rows = self.resolver.rows_from(c["from"])
            except RefError as e:
                self.errors.append(u"%s：chart.from → %s" % (where, e))
                return None
            cat = c.get("category")
            vals = c.get("values") or []
            if isinstance(vals, str):
                vals = [vals]
            missing = [x for x in [cat] + list(vals) if x not in t.columns]
            if not cat or not vals or missing:
                self.errors.append(u"%s：chart 需要 category 與 values，且欄位要存在（缺：%s；%s 欄位：%s）"
                                   % (where, "、".join(str(m) for m in missing) or "category/values", t.id, "、".join(t.columns)))
                return None
            agg = str(c.get("agg", "sum")).lower()
            groups = OrderedDict()
            for r in rows:
                if r.get(cat) is None:
                    continue
                groups.setdefault(format_value(r.get(cat), t.kinds.get(cat)), []).append(r)
            cats = list(groups)
            for col in vals:
                values = []
                for k in cats:
                    nums = [r.get(col) for r in groups[k] if _is_num(r.get(col))]
                    try:
                        values.append(aggregate(nums, "mean" if agg == "avg" else agg))
                    except RefError as e:
                        self.errors.append(u"%s：%s" % (where, e))
                        return None
                series.append({"name": col, "values": values, "percent": t.kinds.get(col) == "percent"})
            if c.get("sort"):
                key = str(c["sort"]).lstrip("-+").strip()
                target = next((s for s in series if s["name"] == key), None)
                if target is None:
                    self.warnings.append(u"%s：chart.sort 的「%s」不在 values 裡，已忽略" % (where, key))
                else:
                    sign = -1 if str(c["sort"]).startswith("-") else 1
                    tv = target["values"]
                    order = sorted(range(len(cats)), key=lambda i: (tv[i] is None, sign * (tv[i] or 0)))
                    cats = [cats[i] for i in order]
                    for s in series:
                        s["values"] = [s["values"][i] for i in order]
            if c.get("limit"):
                n = int(c["limit"])
                cats = cats[:n]
                for s in series:
                    s["values"] = s["values"][:n]
        else:
            cats = [self.txt(x, where) for x in c.get("categories") or []]
            for s in c.get("series") or []:
                kinds = []
                vals = [self.resolver.number(v, self.errors, where, kinds) for v in s.get("values") or []]
                if len(vals) != len(cats):
                    self.errors.append(u"%s：數列「%s」有 %d 個值，但 categories 有 %d 個"
                                       % (where, s.get("name", ""), len(vals), len(cats)))
                series.append({"name": self.txt(s.get("name", ""), where) or "", "values": vals,
                               "percent": bool(s.get("percent")) or (bool(kinds) and all(k == "percent" for k in kinds))})
        if not cats or not series:
            self.errors.append(u"%s：圖表沒有資料（請給 from+category+values，或 categories+series）" % where)
            return None
        if ctype in ("pie", "doughnut") and len(cats) > 5:
            self.warnings.append(u"%s：圓餅圖有 %d 個類別，超過 5 個很難比較，建議改用 bar（橫條）並排序" % (where, len(cats)))
        if len(series) > 4:
            self.warnings.append(u"%s：圖表有 %d 組數列，建議 3 組以內，或拆成兩張圖" % (where, len(series)))
        hl = c.get("highlight") or []
        if isinstance(hl, str):
            hl = [hl]
        hl = [self.txt(h, where) for h in hl]
        unknown = [h for h in hl if h not in [str(x) for x in cats]]
        if unknown:
            self.warnings.append(u"%s：highlight 的「%s」不在類別裡（類別：%s）" % (where, u"、".join(unknown), u"、".join(str(x) for x in cats[:8])))
        return {"type": "chart", "chart_type": ctype, "categories": cats, "series": series,
                "title": self.txt(c.get("title"), where), "labels": bool(c.get("labels", False)),
                "highlight": hl, "highlight_color": c.get("highlight_color")}

    def table(self, tb, where):
        if not isinstance(tb, dict):
            self.errors.append(u"%s：table 必須是物件" % where)
            return None
        if tb.get("from"):
            try:
                t, rows = self.resolver.rows_from(tb["from"])
            except RefError as e:
                self.errors.append(u"%s：table.from → %s" % (where, e))
                return None
            cols = tb.get("columns") or t.columns
            missing = [c for c in cols if c not in t.columns]
            if missing:
                self.errors.append(u"%s：%s 沒有欄位 %s（欄位：%s）" % (where, t.id, "、".join(missing), "、".join(t.columns)))
                return None
            if tb.get("sort"):
                rows = _sort_rows(rows, str(tb["sort"]))
            if tb.get("limit"):
                rows = rows[:int(tb["limit"])]
            fmts = tb.get("formats") or {}
            cells = [[format_value(r.get(c), t.kinds.get(c), fmts.get(c)) for c in cols] for r in rows]
            hl = set()
            h = tb.get("highlight") or {}
            if h.get("column"):
                for i, r in enumerate(rows):
                    v = r.get(h["column"])
                    s = "" if v is None else format_value(v, t.kinds.get(h["column"]))
                    if (h.get("equals") is not None and s == str(h["equals"])) or (h.get("contains") and h["contains"] in s):
                        hl.add(i)
            return {"type": "table", "columns": [self._terms(c) for c in cols], "rows": cells, "highlight": hl,
                    "title": self.txt(tb.get("title"), where)}
        cols = [self.txt(c, where) for c in tb.get("columns") or []]
        cells = [[self.txt(x, where) for x in row] for row in tb.get("rows") or []]
        if not cols:
            self.errors.append(u"%s：table 需要 from，或 columns+rows" % where)
            return None
        return {"type": "table", "columns": cols, "rows": cells, "title": self.txt(tb.get("title"), where),
                "highlight": set(int(i) for i in tb.get("highlight_rows") or [])}

    IMAGE_STYLES = ("card", "browser", "shadow", "plain")

    def image(self, img, where):
        spec = img if isinstance(img, dict) else {"path": img}
        path = spec.get("path")
        p = resolve_path(path) if path else None
        if p is None or not p.exists():
            self.warnings.append(u"%s：找不到圖片「%s」，先放文字占位" % (where, path))
            return {"type": "note", "text": u"[圖片：%s]" % (path or "")}
        if p.suffix.lower() == ".svg":  # vector art becomes editable shapes
            return self.svg({"path": str(p), "title": spec.get("title")}, where)
        style = str(spec.get("style") or "card").lower()
        if style not in self.IMAGE_STYLES:
            self.warnings.append(u"%s：圖片 style「%s」不認得（可用 card/browser/shadow/plain），改用 card" % (where, style))
            style = "card"
        out = {"type": "image", "path": str(p), "style": style, "caption": self.txt(spec.get("caption"), where),
               "url": spec.get("url"), "fit": spec.get("fit"), "trim": spec.get("trim", True) is not False}
        crop = spec.get("crop")
        if crop is not None:
            if isinstance(crop, dict) or (isinstance(crop, list) and len(crop) == 4):
                out["crop"] = crop
            else:
                self.warnings.append(u"%s：crop 要是 [左, 上, 右, 下] 或 {x,y,w,h}，已忽略" % where)
        hls = []
        for i, hl in enumerate(spec.get("highlights") or [], 1):
            if not isinstance(hl, dict) or not all(k in hl for k in ("x", "y", "w", "h")):
                self.warnings.append(u"%s：highlights 第 %d 項需要 x,y,w,h，已略過" % (where, i))
                continue
            hl = dict(hl)
            if hl.get("note"):
                hl["note"] = self._terms(self.txt(hl["note"], where))
            hls.append(hl)
        out["highlights"] = hls
        z = spec.get("zoom")
        if isinstance(z, dict) and all(k in z for k in ("x", "y", "w", "h")):
            out["zoom"] = z
        elif z:
            self.warnings.append(u"%s：zoom 需要 x,y,w,h，已忽略" % where)
        return out

    def images(self, imgs, where):
        spec = imgs if isinstance(imgs, dict) else {"items": imgs}
        items = []
        for it in spec.get("items") or []:
            b = self.image(it, where)
            if b and b["type"] == "image":
                items.append(b)
        if not items:
            self.errors.append(u"%s：圖組沒有可用的圖片" % where)
            return None
        if len(items) > 6:
            self.warnings.append(u"%s：一頁 %d 張圖偏多，建議 6 張內（拆頁或只留重點）" % (where, len(items)))
        return {"type": "images", "items": items, "columns": spec.get("columns"), "title": self.txt(spec.get("title"), where)}

    # -- drawings -------------------------------------------------------
    def _date(self, v, where, what):
        if v is None or v == "":
            return None
        if isinstance(v, str) and str(v).strip().lower() in ("today", u"今天"):
            return _dt.date.today()
        if isinstance(v, str) and "{" in v:
            v = self.resolver.render(v, self.errors, where)
        try:
            return to_date(v)
        except ValueError as e:
            self.errors.append(u"%s：%s %s" % (where, what, e))
            return None

    def _ratio(self, v, where):
        if v is None or v == "":
            return None
        n = self.resolver.number(v, self.errors, where)
        if n is None:
            return None
        return n / 100.0 if n > 1.0 else n

    def _rows_or_items(self, spec, key, cols, where, label):
        """Items from spec[key], or from a table: {"from": "t1", <field>: <column>…}."""
        if isinstance(spec, list):
            return spec
        if not isinstance(spec, dict):
            self.errors.append(u"%s：%s 必須是物件或陣列" % (where, label))
            return None
        if not spec.get("from"):
            return spec.get(key) or []
        try:
            t, rows = self.resolver.rows_from(spec["from"])
        except RefError as e:
            self.errors.append(u"%s：%s.from → %s" % (where, label, e))
            return None
        mapping = {k: spec.get(k) for k in cols if spec.get(k)}
        missing = [c for c in mapping.values() if c not in t.columns]
        if missing:
            self.errors.append(u"%s：%s 沒有欄位 %s（欄位：%s）" % (where, t.id, u"、".join(missing), u"、".join(t.columns)))
            return None
        if spec.get("sort"):
            rows = _sort_rows(rows, str(spec["sort"]))
        out = []
        for r in rows[:int(spec.get("limit") or len(rows))]:
            it = {}
            for k, c in mapping.items():
                v = r.get(c)
                it[k] = v if (_is_num(v) or v is None) else format_value(v, t.kinds.get(c))
            out.append(it)
        return out

    def gantt(self, g, where):
        items = self._rows_or_items(g, "tasks", ("name", "start", "end", "progress", "status", "group", "owner"),
                                    where, "gantt")
        if items is None:
            return None
        g = g if isinstance(g, dict) else {}
        tasks = []
        for i, it in enumerate(items, 1):
            if isinstance(it, str):
                self.errors.append(u"%s：甘特圖第 %d 項需要 name、start、end" % (where, i))
                continue
            w = u"%s 甘特圖「%s」" % (where, it.get("name") or i)
            start = self._date(it.get("start"), w, "start")
            end = self._date(it.get("end"), w, "end") or start
            if start is None:
                if it.get("start") in (None, ""):
                    self.errors.append(u"%s：缺少 start" % w)
                continue
            if end < start:
                self.warnings.append(u"%s：end 早於 start，已對調" % w)
                start, end = end, start
            tasks.append({"name": self._terms(self.txt(it.get("name"), where) or u"—"), "start": start, "end": end,
                          "progress": self._ratio(it.get("progress"), w),
                          "status": status_key(it.get("status")) if it.get("status") else None,
                          "group": self.txt(it.get("group"), where) or "", "owner": self.txt(it.get("owner"), where) or "",
                          "milestone": bool(it.get("milestone"))})
        miles = []
        for m in g.get("milestones") or []:
            d = self._date(m.get("date"), where, u"milestones.date")
            if d:
                miles.append({"name": self.txt(m.get("name"), where) or "", "date": d,
                              "status": status_key(m.get("status")) if m.get("status") else None})
        if not tasks and not miles:
            if not items:  # otherwise each bad item has already been reported
                self.errors.append(u"%s：甘特圖沒有任何工作項" % where)
            return None
        if len(tasks) + len(miles) > 14:
            self.warnings.append(u"%s：甘特圖 %d 列偏多，建議 12 列內（用 group 分組或拆頁）" % (where, len(tasks) + len(miles)))
        return {"type": "gantt", "tasks": tasks, "milestones": miles, "today": self._date(g.get("today"), where, "today"),
                "title": self.txt(g.get("title"), where)}

    def timeline(self, tl, where):
        items = self._rows_or_items(tl, "items", ("date", "label", "note", "status", "state", "when"), where, "timeline")
        if items is None:
            return None
        tl = tl if isinstance(tl, dict) else {}
        out = []
        for it in items:
            if isinstance(it, str):
                it = {"label": it}
            state = str(it.get("state") or "").lower()
            state = {"done": "done", u"完成": "done", u"已完成": "done", "current": "current", u"進行中": "current",
                     "now": "current", "todo": "todo", u"未開始": "todo"}.get(state)
            out.append({"date": self._date(it.get("date"), where, "date"), "label": self._terms(self.txt(it.get("label"), where) or ""),
                        "note": self._terms(self.txt(it.get("note"), where) or ""), "when": self.txt(it.get("when"), where),
                        "state": state, "status": status_key(it.get("status")) if it.get("status") else None})
        if not out:
            self.errors.append(u"%s：時程表沒有項目" % where)
            return None
        if len(out) > 8:
            self.warnings.append(u"%s：時程表 %d 個節點偏多，建議 8 個內" % (where, len(out)))
        return {"type": "timeline", "items": out, "today": self._date(tl.get("today"), where, "today"),
                "even": bool(tl.get("even")), "title": self.txt(tl.get("title"), where)}

    def process(self, pr, where):
        items = self._rows_or_items(pr, "steps", ("title", "note", "owner", "state", "status"), where, "process")
        if items is None:
            return None
        pr = pr if isinstance(pr, dict) else {}
        steps = []
        for it in items:
            if isinstance(it, str):
                it = {"title": it}
            state = str(it.get("state") or "").lower()
            state = {"done": "done", u"完成": "done", "current": "current", u"進行中": "current",
                     "todo": "todo", u"未開始": "todo"}.get(state)
            steps.append({"title": self._terms(self.txt(it.get("title"), where) or ""), "note": self._terms(self.txt(it.get("note"), where) or ""),
                          "owner": self.txt(it.get("owner"), where) or "", "state": state,
                          "status": status_key(it.get("status")) if it.get("status") else None})
        if not steps:
            self.errors.append(u"%s：流程沒有步驟" % where)
            return None
        if len(steps) > 6:
            self.warnings.append(u"%s：流程 %d 步偏多，建議 6 步內（合併相近步驟）" % (where, len(steps)))
        return {"type": "process", "steps": steps, "title": self.txt(pr.get("title"), where)}

    def svg(self, sv, where):
        title = None
        if isinstance(sv, dict):
            title = sv.get("title")
            if sv.get("path"):
                p = resolve_path(sv["path"])
                if not p.exists():
                    self.errors.append(u"%s：找不到 SVG 檔「%s」" % (where, sv["path"]))
                    return None
                markup = p.read_text(encoding="utf-8")
            else:
                markup = sv.get("svg") or sv.get("markup") or ""
        else:
            markup = sv
        try:
            root = parse_svg(markup)
        except ValueError as e:
            self.errors.append(u"%s：%s" % (where, e))
            return None
        for el in root.iter():  # {references} and term rules apply to the drawing's text, not its coordinates
            if local(el.tag) in ("text", "tspan"):
                if el.text and el.text.strip():
                    el.text = self._terms(self.resolver.render(el.text, self.errors, where))
                if el.tail and el.tail.strip():
                    el.tail = self._terms(self.resolver.render(el.tail, self.errors, where))
        return {"type": "svg", "svg": etree.tostring(root, encoding="unicode"), "title": self.txt(title, where)}

    def region(self, r, where):
        if not isinstance(r, dict):
            self.errors.append(u"%s：region 必須是物件" % where)
            return None
        raw_type = str(r.get("type", "")).strip()
        rtype = REGION_TYPES.get(raw_type.lower()) or REGION_TYPES.get(raw_type)
        if not rtype:
            self.errors.append(u"%s：不認得區塊類型「%s」（可用：title/headline/bullets/text/kpi/chart/table/status/"
                               u"gantt/timeline/process/svg/image/note）"
                               % (where, raw_type))
            return None
        try:
            x, y, w, h = (float(r[k]) for k in ("x", "y", "w", "h"))
        except (KeyError, TypeError, ValueError):
            self.errors.append(u"%s：需要 x,y,w,h 格線座標（12×8）" % where)
            return None
        x, y = min(max(x, 0.0), 11.5), min(max(y, 0.0), 7.5)
        w, h = max(0.5, min(w, 12.0 - x)), max(0.5, min(h, 8.0 - y))
        content = r.get("content")
        if content is None:
            for k in ("text", "bullets", "kpis", "chart", "table", "image", "status") + DIAGRAM_KEYS + ("items", "value"):
                if k in r:
                    content = r[k]
                    break
        if content is None or content == "" or content == []:
            self.warnings.append(u"%s：「%s」區塊沒有內容，已留白" % (where, raw_type))
            return None
        if rtype == "title":
            return {"type": "title", "text": self.txt(content if isinstance(content, str) else str(content), where)}
        if rtype == "headline":
            block = {"type": "headline", "text": self.txt(str(content), where)}
        elif rtype == "bullets":
            block = {"type": "text", "paras": self.paras(content, where)}
        elif rtype == "text":
            block = {"type": "text", "paras": self.plain_paras(content, where)}
        elif rtype == "kpi":
            block = self.kpis(content, where)
        elif rtype == "chart":
            block = self.chart(content, where)
        elif rtype == "table":
            block = self.table(content, where)
        elif rtype == "image":
            block = self.image(content, where)
        elif rtype == "status":
            block = self.status(content, where)
        elif rtype in DIAGRAM_KEYS:
            block = getattr(self, rtype)(content, where)
        else:
            block = {"type": "note", "text": self.txt(str(content), where)}
        if block is None:
            return None
        return {"rect": (x, y, w, h), "type": rtype, "block": block}

    # -- slides ---------------------------------------------------------
    def _blueprint_regions(self, layout, fill, where):
        name, _, num = layout.partition("/")
        pages = self.blueprints[name].get("slides") or []
        try:
            page = pages[int(num) - 1]
        except (ValueError, IndexError):
            self.errors.append(u"%s：已存版型「%s」沒有第 %s 頁（共 %d 頁）" % (where, name, num, len(pages)))
            return []
        regs = []
        for r in page.get("regions") or []:
            rr = dict(r)
            if r.get("id") in fill:
                rr["content"] = fill[r["id"]]
            regs.append(rr)
        return regs

    def normalize(self, s, idx, label=None):
        where = u"第 %d 頁%s" % (idx, u"（%s）" % label if label else "")
        ns = {"idx": idx, "label": label, "where": where, "page": s.get("source_page"),
              "title": self.txt(s.get("title"), where), "notes": self.txt(s.get("notes"), where),
              "source": self.txt(s.get("source"), where)}
        if str(s.get("type", "")).lower() in ("section", u"章節", u"章節頁"):
            ns["mode"] = "section"
            ns["subtitle"] = self.txt(s.get("subtitle") or s.get("headline"), where)
            if not ns["title"]:
                self.errors.append(u"%s：章節頁需要 title" % where)
            return ns
        layout = s.get("layout")
        regions_spec = None
        if isinstance(layout, str) and "/" in layout and layout.split("/")[0] in self.blueprints:
            regions_spec = self._blueprint_regions(layout, s.get("fill") or {}, where)
        elif layout == "custom" or (not layout and s.get("regions")):
            regions_spec = s.get("regions") or []
        bullets = []
        if regions_spec is not None:
            ns["mode"], ns["regions"] = "custom", []
            for j, r in enumerate(regions_spec, 1):
                reg = self.region(r, u"%s 區塊 %d" % (where, j))
                if reg is None:
                    continue
                if reg["type"] == "title":
                    ns["title"] = ns["title"] or reg["text"]
                    continue
                ns["regions"].append(reg)
                if reg["block"]["type"] == "text":
                    bullets.extend(p for p in reg["block"]["paras"] if p.get("bullet"))
            headline = any(r["type"] == "headline" for r in ns["regions"])
        else:
            ns["mode"], ns["layout"] = "layout", layout
            headline = self.txt(s.get("headline"), where)
            paras = [{"text": headline, "level": 0, "emph": None, "bullet": False, "headline": True}] if headline else []
            bullets = self.paras(s.get("bullets"), where)
            paras += bullets
            if s.get("text"):
                paras += self.plain_paras(s["text"], where)
            blocks = [{"type": "text", "paras": paras}] if paras else []
            for key, fn in (("kpis", self.kpis), ("status", self.status), ("chart", self.chart), ("table", self.table),
                            ("gantt", self.gantt), ("timeline", self.timeline), ("process", self.process),
                            ("svg", self.svg), ("images", self.images), ("image", self.image)):
                if s.get(key):
                    b = fn(s[key], where)
                    if b:
                        blocks.append(b)
            ns["blocks"] = blocks
        self._audience_checks(ns, bullets, bool(headline), where)
        return ns

    def _audience_checks(self, ns, bullets, has_headline, where):
        a = self.aud
        mb, mc = a.get("max_bullets_per_slide"), a.get("max_chars_per_bullet")
        if mb and len(bullets) > int(mb):
            self.warnings.append(u"%s：%d 個條列，超過「%s」建議的 %s 個" % (where, len(bullets), a.get("name", self.aud_key), mb))
        if mc:
            long_ = [plain_text(p["text"]) for p in bullets if len(plain_text(p["text"])) > int(mc)]
            if long_:
                self.warnings.append(u"%s：%d 個條列超過 %s 字（如「%s…」）" % (where, len(long_), mc, long_[0][:12]))
        if a.get("require_headline") and not ns.get("title") and not has_headline:
            self.warnings.append(u"%s：沒有標題或結論句" % where)

    def expand_repeat(self, s, idx):
        rep = s.get("repeat") or {}
        where = u"第 %d 頁 repeat" % idx
        over, by = rep.get("over"), rep.get("by")
        if not over or not by:
            self.errors.append(u"%s：repeat 需要 over（例如 t1）與 by（欄位）" % where)
            return []
        try:
            t, rows = self.resolver.rows_from(over)
        except RefError as e:
            self.errors.append(u"%s：%s" % (where, e))
            return []
        if by not in t.columns:
            self.errors.append(u"%s：%s 沒有欄位「%s」（欄位：%s）" % (where, t.id, by, "、".join(t.columns)))
            return []
        keys = []
        for r in rows:
            if r.get(by) is None:
                continue
            k = format_value(r.get(by), t.kinds.get(by))
            if k not in keys:
                keys.append(k)
        if not keys:
            self.warnings.append(u"%s：沒有符合的資料列，這頁已略過" % where)
        if len(keys) > MAX_REPEAT:
            self.warnings.append(u"%s：展開 %d 頁太多，只做前 %d 頁" % (where, len(keys), MAX_REPEAT))
            keys = keys[:MAX_REPEAT]
        base = {k: v for k, v in s.items() if k != "repeat"}
        return [(_deep_replace(base, {"[@]": "[%s=%s]" % (by, k), "{$key}": k}), k) for k in keys]


def _save(renderer, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    cand = path
    for n in range(2, 30):
        try:
            renderer.save(cand)
            return cand
        except PermissionError:  # file is open in PowerPoint
            cand = path.with_name("%s (%d)%s" % (path.stem, n, path.suffix))
    raise PermissionError("無法寫入 %s（檔案可能正在 PowerPoint 開啟）" % path)


def build(spec, out_path=None, base_dir=None) -> dict:
    if not isinstance(spec, dict) or not isinstance(spec.get("slides"), list) or not spec["slides"]:
        return {"ok": False, "errors": [u"spec 需要是物件，且 slides 至少一頁"], "warnings": []}
    b = Build(spec, base_dir)
    nslides = []
    for i, s in enumerate(spec["slides"], 1):
        if not isinstance(s, dict):
            b.errors.append(u"第 %d 頁：必須是物件" % i)
            continue
        if s.get("page_type") in ("requirements", u"需求"):
            continue
        if s.get("repeat"):
            for s2, key in b.expand_repeat(s, i):
                nslides.append(b.normalize(s2, i, key))
        else:
            nslides.append(b.normalize(s, i))
    ms = b.aud.get("max_slides")
    if ms and len(nslides) > int(ms):
        b.warnings.append(u"共 %d 頁，超過「%s」建議的 %s 頁" % (len(nslides), b.aud.get("name", b.aud_key), ms))
    if b.errors:
        return {"ok": False, "errors": b.errors, "warnings": b.warnings}

    template = resolve_template_path(spec.get("template"), b.cfg)
    r = DeckRenderer(template, b.cfg)
    r.meta = {k: b.txt(spec.get(k), u"封面") for k in ("author", "date") if spec.get(k)}
    produced = r.render_all(nslides)
    b.warnings.extend(r.warnings)

    if out_path:
        target = resolve_path(out_path, base_dir)
        if target.suffix.lower() != ".pptx":
            target = target / ("%s_%s.pptx" % (_safe_name(spec.get("title") or "report"), b.aud_key))
    else:
        odir = Path(b.cfg["output_dir"]) if b.cfg.get("output_dir") else output_dir()
        target = odir / ("%s_%s.pptx" % (_safe_name(spec.get("title") or "report"), b.aud_key))
    saved = _save(r, target)

    by_idx = OrderedDict()
    for ns, nums in zip(nslides, produced):
        by_idx.setdefault(ns["idx"], []).append((ns.get("label"), nums))
    mapping = []
    for idx, parts in by_idx.items():
        nums = [n for _l, ns_ in parts for n in ns_]
        labels = [l for l, _n in parts if l]
        mapping.append({"spec": idx, "slides": nums, "repeat": labels})
    slides_of = {m["spec"]: m["slides"] for m in mapping}
    annotations = []
    for a in spec.get("annotations") or []:
        a = a if isinstance(a, dict) else {"raw": str(a)}
        annotations.append({"page": a.get("page", u"需求"), "slides": [], "raw": a.get("raw", ""), "applied": a.get("applied", "")})
    for i, s in enumerate(spec["slides"], 1):
        if not isinstance(s, dict):
            continue
        for a in s.get("annotations") or []:
            a = a if isinstance(a, dict) else {"raw": str(a)}
            annotations.append({"page": s.get("source_page", ""), "slides": slides_of.get(i, []),
                                "raw": a.get("raw", ""), "applied": a.get("applied", "")})

    saved_spec = dict(spec)
    saved_spec["sources"] = [str(resolve_path(p, base_dir)) for p in spec.get("sources") or []]
    saved_spec["_output"] = str(saved)
    spec_path = saved.with_suffix(".spec.json")
    C.write_json(spec_path, saved_spec)
    if b.replaced:
        b.warnings.append(u"已自動替換用語 %d 處（style_rules.yaml）" % b.replaced)
    try:
        from .lint import lint_deck
        lint = lint_deck(saved, b.cfg)
    except Exception as e:  # lint must never block delivering the deck
        lint = {"slide_count": len(r.prs.slides), "issues": [{"slide": 0, "level": "info", "msg": u"版面檢查失敗：%s" % e}]}
    return {"ok": True, "lint": lint, "path": str(saved), "spec_path": str(spec_path), "slide_count": len(r.prs.slides),
            "template": str(template), "audience": b.aud_key, "mapping": mapping,
            "annotations": annotations, "warnings": b.warnings}


def patch(spec_path, slides=None, delete=None, insert=None, out_path=None) -> dict:
    p = resolve_path(spec_path)
    spec = C.read_json(p, None)
    if not isinstance(spec, dict):
        return {"ok": False, "errors": [u"讀不到 spec：%s" % spec_path], "warnings": []}
    orig = spec.get("slides") or []
    errors = []
    for k, change in (slides or {}).items():
        try:
            i = int(k) - 1
        except ValueError:
            errors.append(u"slides 的 key 要是頁碼：%s" % k)
            continue
        if not 0 <= i < len(orig) or not isinstance(change, dict):
            errors.append(u"沒有第 %s 頁（共 %d 頁），或內容不是物件" % (k, len(orig)))
            continue
        if change.get("_replace"):
            orig[i] = {kk: vv for kk, vv in change.items() if kk != "_replace"}
        else:
            new = dict(orig[i])
            for kk, vv in change.items():
                if vv is None:
                    new.pop(kk, None)
                else:
                    new[kk] = vv
            orig[i] = new
    if errors:
        return {"ok": False, "errors": errors, "warnings": []}
    slots = [[s] for s in orig]
    head = []
    for d in delete or []:
        if 1 <= int(d) <= len(slots):
            slots[int(d) - 1] = []
    for ins in insert or []:
        after = int(ins.get("after", len(slots)))
        (head if after <= 0 else slots[min(after, len(slots)) - 1]).append(ins.get("slide") or {})
    spec["slides"] = head + [s for slot in slots for s in slot]
    out = out_path or spec.pop("_output", None)
    spec.pop("_output", None)
    return build(spec, out, base_dir=p.parent)
