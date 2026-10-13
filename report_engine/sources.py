"""Load report data: Excel/CSV become tables (t1, t2 …); Word/PowerPoint/PDF give text (d1 …) plus their tables."""
from __future__ import annotations

import copy
import csv
import datetime as _dt
import io
import re
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Dict, List

from .paths import resolve_path

TABLE_EXT = {".xlsx", ".xlsm", ".csv", ".tsv"}
DOC_EXT = {".docx", ".pptx", ".pdf", ".md", ".txt"}
SUPPORTED = TABLE_EXT | DOC_EXT
MAX_ROWS = 20000
MAX_PDF_PAGES = 60

_NUM_RE = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")
_PCT_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?\s*%$")
_DATE_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$")


class Pct(float):
    """A number Excel displays as a percentage."""


class Table:
    def __init__(self, source: str, sheet: str, columns: List[str], rows: List[dict], kinds: Dict[str, str]):
        self.id = ""
        self.source = source
        self.sheet = sheet
        self.columns = columns
        self.rows = rows
        self.kinds = kinds

    @property
    def label(self) -> str:
        name = Path(self.source).name
        return "%s › %s" % (name, self.sheet) if self.sheet else name


class Document:
    def __init__(self, source: str, text: str):
        self.id = ""
        self.source = source
        self.text = text


class Bundle:
    def __init__(self):
        self.tables = OrderedDict()  # id -> Table
        self.documents = []  # List[Document]
        self.errors = []  # List[str]
        self.files = []  # List[str]


def _nonempty(c) -> bool:
    return c is not None and str(c).strip() != ""


def _convert(v):
    """Return (value, kind) with kind in number/percent/date/text, or (None, None) for empty."""
    if v is None:
        return None, None
    if isinstance(v, Pct):
        return float(v), "percent"
    if isinstance(v, bool):
        return v, "text"
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
            return int(v), "number"
        return v, "number"
    if isinstance(v, _dt.datetime):
        return (v.date().isoformat() if v.time() == _dt.time(0) else v.strftime("%Y-%m-%d %H:%M")), "date"
    if isinstance(v, _dt.date):
        return v.isoformat(), "date"
    if isinstance(v, _dt.time):
        return v.strftime("%H:%M"), "text"
    s = str(v).strip()
    if not s:
        return None, None
    if _PCT_RE.match(s):
        return float(s.rstrip("%").strip()) / 100.0, "percent"
    if _NUM_RE.match(s) and not (len(s) > 1 and s[0] == "0" and s[1] != "."):
        t = s.replace(",", "")
        return (float(t) if "." in t else int(t)), "number"
    if _DATE_RE.match(s):
        return s.replace("/", "-"), "date"
    return s, "text"


def _kind(c: Counter) -> str:
    total = sum(c.values())
    if not total:
        return "text"
    if c["number"] + c["percent"] >= 0.8 * total:
        return "percent" if c["percent"] >= c["number"] else "number"
    if c["date"] >= 0.8 * total:
        return "date"
    return "text"


def _header_name(h, j) -> str:
    if isinstance(h, float) and h.is_integer():
        h = int(h)
    if isinstance(h, (_dt.datetime, _dt.date)):
        h = h.strftime("%Y-%m-%d")
    return str(h).strip().replace("\n", " ") if _nonempty(h) else "欄%d" % (j + 1)


def build_table(source: str, sheet: str, raw) -> "Table":
    rows = [list(r) for r in raw if any(_nonempty(c) for c in r)]
    if len(rows) < 2:
        return None
    head = rows[:10]
    widest = max(sum(1 for c in r if _nonempty(c)) for r in head)
    hi = next(i for i, r in enumerate(head) if sum(1 for c in r if _nonempty(c)) >= max(1, widest * 0.5))
    width = max(len(r) for r in rows)
    header = rows[hi] + [None] * (width - len(rows[hi]))
    data = [r + [None] * (width - len(r)) for r in rows[hi + 1:]]
    if not data:
        return None
    keep = [j for j in range(width) if _nonempty(header[j]) or any(_nonempty(d[j]) for d in data)]
    names, seen = [], {}
    for j in keep:
        n = _header_name(header[j], j)
        if n in seen:
            seen[n] += 1
            n = "%s_%d" % (n, seen[n])
        else:
            seen[n] = 1
        names.append(n)
    counts = {n: Counter() for n in names}
    out = []
    for d in data:
        rec = {}
        for n, j in zip(names, keep):
            v, k = _convert(d[j])
            rec[n] = v
            if k:
                counts[n][k] += 1
        out.append(rec)
    return Table(source, sheet, names, out, {n: _kind(counts[n]) for n in names})


# ---- readers -------------------------------------------------------------

def _read_xlsx(path: Path):
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            if getattr(ws, "sheet_state", "visible") != "visible":
                continue
            raw = []
            for i, row in enumerate(ws.iter_rows()):
                if i >= MAX_ROWS:
                    break
                vals = []
                for c in row:
                    v = getattr(c, "value", None)
                    fmt = getattr(c, "number_format", "") or ""
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and "%" in fmt:
                        v = Pct(v)
                    vals.append(v)
                raw.append(vals)
            out.append((ws.title, raw))
    finally:
        wb.close()
    return out


def _decode(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    for enc in ("utf-8-sig", "cp950", "big5hkscs"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _read_csv(path: Path):
    text = _decode(path.read_bytes())
    delim = "\t" if path.suffix.lower() == ".tsv" else None
    if not delim:
        try:
            delim = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
        except csv.Error:
            delim = ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))[:MAX_ROWS]
    return [("", rows)]


def _iter_shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            for x in _iter_shapes(s.shapes):
                yield x
        else:
            yield s


def _chart_rows(chart):
    try:
        plot = chart.plots[0]
        cats = list(plot.categories)
        series = list(plot.series)
        rows = [["類別"] + [s.name for s in series]]
        for i, c in enumerate(cats):
            rows.append([str(c)] + [(s.values[i] if i < len(s.values) else None) for s in series])
        return rows
    except Exception:
        return None


def _read_pptx(path: Path):
    from pptx import Presentation

    prs = Presentation(str(path))
    texts, tables = [], []
    for n, slide in enumerate(prs.slides, 1):
        parts = []
        for shp in _iter_shapes(slide.shapes):
            if shp.has_text_frame and shp.text_frame.text.strip():
                parts.append(shp.text_frame.text.strip())
            if getattr(shp, "has_table", False) and shp.has_table:
                tables.append(("P%d 表格" % n, [[c.text for c in r.cells] for r in shp.table.rows]))
            if getattr(shp, "has_chart", False) and shp.has_chart:
                rows = _chart_rows(shp.chart)
                if rows:
                    tables.append(("P%d 圖表" % n, rows))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                parts.append("（備註）" + notes)
        texts.append("--- 第 %d 頁 ---\n%s" % (n, "\n".join(parts)))
    return "\n".join(texts), tables


def _read_docx(path: Path):
    import docx

    d = docx.Document(str(path))
    parts = []
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        style = (p.style.name or "") if p.style is not None else ""
        if style.lower().startswith("heading") or style.startswith("標題"):
            m = re.search(r"\d", style)
            parts.append("#" * (int(m.group()) if m else 1) + " " + t)
        else:
            parts.append(t)
    tables = [("表格%d" % i, [[c.text for c in r.cells] for r in t.rows]) for i, t in enumerate(d.tables, 1)]
    return "\n".join(parts), tables


def _read_pdf(path: Path):
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(path))
    try:
        out = []
        for i in range(min(len(pdf), MAX_PDF_PAGES)):
            page = pdf[i]
            tp = page.get_textpage()
            out.append("--- 第 %d 頁 ---\n%s" % (i + 1, tp.get_text_range().strip()))
            tp.close()
            page.close()
        return "\n".join(out), []
    finally:
        pdf.close()


_CACHE = {}


def _parse_file(path: Path):
    st = path.stat()
    key = (str(path), st.st_mtime_ns, st.st_size)
    if key in _CACHE:
        return _CACHE[key]
    ext = path.suffix.lower()
    text, raw_tables = None, []
    if ext in (".xlsx", ".xlsm"):
        raw_tables = _read_xlsx(path)
    elif ext in (".csv", ".tsv"):
        raw_tables = _read_csv(path)
    elif ext == ".pptx":
        text, raw_tables = _read_pptx(path)
    elif ext == ".docx":
        text, raw_tables = _read_docx(path)
    elif ext == ".pdf":
        text, raw_tables = _read_pdf(path)
    else:
        text = _decode(path.read_bytes())
    tables = [t for t in (build_table(str(path), sheet, raw) for sheet, raw in raw_tables) if t is not None]
    _CACHE[key] = (tables, text)
    return _CACHE[key]


def load_sources(paths, base=None) -> Bundle:
    b = Bundle()
    files = []
    for p in paths or []:
        path = resolve_path(p, base)
        if path.is_dir():
            files.extend(sorted(x for x in path.iterdir()
                                if x.suffix.lower() in SUPPORTED and not x.name.startswith("~$")))
        elif path.exists():
            files.append(path)
        else:
            b.errors.append("找不到資料來源：%s" % p)
    for f in files:
        ext = f.suffix.lower()
        if ext == ".xls":
            b.errors.append("%s 是舊版 .xls，請另存為 .xlsx" % f.name)
            continue
        if ext not in SUPPORTED:
            b.errors.append("不支援的檔案類型：%s" % f.name)
            continue
        try:
            tables, text = _parse_file(f)
        except Exception as e:  # report and keep going with the other files
            b.errors.append("讀取 %s 失敗：%s: %s" % (f.name, type(e).__name__, e))
            continue
        b.files.append(str(f))
        for t in tables:
            tc = copy.copy(t)
            tc.id = "t%d" % (len(b.tables) + 1)
            b.tables[tc.id] = tc
        if text and text.strip():
            d = Document(str(f), text)
            d.id = "d%d" % (len(b.documents) + 1)
            b.documents.append(d)
    return b
