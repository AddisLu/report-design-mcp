"""Deterministic number references so the model never computes or copies numbers itself.

Syntax inside any text: {t1[專案=B].交付率 | .1%}
  t1                 table id from get_design_context
  [col=val]          filters (= != > < >= <= ~contains), can repeat
  .col               column
  .sum/.mean/...     aggregate
  {=a - b}           arithmetic over references / named facts
  {name}             named fact from spec.facts
  | fmt              Python format spec
"""
from __future__ import annotations

import ast
import math
import re
import statistics

BRACE_RE = re.compile(r"\{([^{}]+)\}")
REF_HEAD = re.compile(r"^(t\d+)")
REF_IN_EXPR = re.compile(r"(?<![\w$])t\d+(?:\[[^\]]*\])*(?:\.[^\s+\-*/()\[\],`|]+)*")
FILTER_RE = re.compile(r"^(.+?)(!=|>=|<=|=|~|>|<)(.*)$", re.S)
AGGS = {"sum", "mean", "avg", "min", "max", "median", "count", "distinct", "first", "last"}


class RefError(Exception):
    pass


def parse_ref(s: str):
    s = s.strip()
    m = REF_HEAD.match(s)
    if not m:
        raise RefError("引用必須以表格 id 開頭，例如 t1")
    tid, pos, filters = m.group(1), m.end(), []
    while pos < len(s) and s[pos] == "[":
        end = s.find("]", pos)
        if end < 0:
            raise RefError("篩選條件缺少 ]")
        cond = s[pos + 1:end]
        if cond.strip() == "@":
            raise RefError("[@] 只能用在有 repeat 的頁面")
        fm = FILTER_RE.match(cond)
        if not fm:
            raise RefError("篩選條件「%s」格式應為 [欄位=值]" % cond)
        filters.append((fm.group(1).strip(), fm.group(2), fm.group(3).strip()))
        pos = end + 1
    rest = s[pos:].strip()
    col = agg = None
    if rest:
        if not rest.startswith("."):
            raise RefError("無法解析「%s」" % rest)
        parts = rest[1:].split(".")
        if len(parts) > 1 and parts[-1].strip().lower() in AGGS:
            agg = parts.pop().strip().lower()
        elif len(parts) == 1 and parts[0].strip().lower() in ("count", "rows"):
            agg, parts = "count", []
        col = ".".join(parts).strip() or None
    if agg == "avg":
        agg = "mean"
    return tid, filters, col, agg


def to_num(s):
    t = str(s).strip().replace(",", "")
    pct = t.endswith("%")
    if pct:
        t = t[:-1].strip()
    try:
        v = float(t)
    except ValueError:
        return None
    return v / 100.0 if pct else v


def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _match(cell, op, val, kind):
    if op == "~":
        return cell is not None and val in str(cell)
    vn = to_num(val)
    if vn is not None and kind == "percent" and not val.strip().endswith("%") and abs(vn) > 1:
        vn = vn / 100.0  # [進度<70] on a percent column means 70%
    cn = cell if _is_num(cell) else None
    if op in (">", "<", ">=", "<="):
        if cn is None or vn is None:
            return False
        return {">": cn > vn, "<": cn < vn, ">=": cn >= vn, "<=": cn <= vn}[op]
    if cn is not None and vn is not None:
        eq = math.isclose(cn, vn, rel_tol=1e-9, abs_tol=1e-12)
    else:
        eq = ("" if cell is None else str(cell).strip()) == val
    return eq if op == "=" else not eq


def filter_rows(table, filters):
    rows = table.rows
    for col, op, val in filters:
        if col not in table.columns:
            raise RefError("%s 沒有欄位「%s」（欄位：%s）" % (table.id, col, "、".join(table.columns)))
        kind = table.kinds.get(col)
        rows = [r for r in rows if _match(r.get(col), op, val, kind)]
    return rows


def aggregate(nums, agg):
    if agg == "count":
        return len(nums)
    if not nums:
        return None
    if agg == "sum":
        return sum(nums)
    if agg == "mean":
        return sum(nums) / float(len(nums))
    if agg == "min":
        return min(nums)
    if agg == "max":
        return max(nums)
    if agg == "median":
        return statistics.median(nums)
    if agg == "first":
        return nums[0]
    if agg == "last":
        return nums[-1]
    raise RefError("不支援的聚合 %s" % agg)


def format_value(v, kind=None, fmt=None) -> str:
    if v is None or (isinstance(v, str) and v == ""):
        return "—"
    if fmt:
        try:
            return format(v, fmt)
        except (ValueError, TypeError):
            return str(v)
    if isinstance(v, bool):
        return "是" if v else "否"
    if _is_num(v):
        if kind == "percent":
            return ("%.1f" % (v * 100)).rstrip("0").rstrip(".") + "%"
        if isinstance(v, int) or (abs(v) < 1e15 and float(v).is_integer()):
            return "{:,}".format(int(v))
        if abs(v) >= 100:
            return "{:,.0f}".format(v)
        return "{:,.2f}".format(v).rstrip("0").rstrip(".")
    return str(v)


def split_fmt(inner: str):
    if "|" in inner:
        body, fmt = inner.rsplit("|", 1)
        return body.strip(), (fmt.strip() or None)
    return inner.strip(), None


_FUNCS = {"abs": abs, "min": min, "max": max, "round": lambda x, n=0: round(x, int(n))}
_BIN = {ast.Add: lambda a, b: a + b, ast.Sub: lambda a, b: a - b,
        ast.Mult: lambda a, b: a * b, ast.Div: lambda a, b: a / b}


class Resolver:
    def __init__(self, tables, named=None):
        self.tables = tables
        self.named = named or {}
        self._named_cache = {}
        self._stack = []

    # -- public ---------------------------------------------------------
    def render(self, text, errors, where="") -> str:
        """Replace every {…} in text with its formatted value; unresolved refs go to errors."""
        if text is None:
            return ""
        if not isinstance(text, str):
            return format_value(text)

        def rep(m):
            inner = m.group(1)
            try:
                body, fmt = split_fmt(inner)
                v, kind = self.evaluate(body)
                return format_value(v, kind, fmt)
            except RefError as e:
                errors.append("%s：{%s} → %s" % (where, inner, e))
                return "{" + inner + "}"

        return BRACE_RE.sub(rep, text)

    def number(self, x, errors, where="", kinds=None):
        if x is None:
            return None
        if _is_num(x):
            return x
        s = str(x).strip()
        m = re.fullmatch(r"\{([^{}]+)\}", s)
        if m:
            body, _fmt = split_fmt(m.group(1))
            try:
                v, k = self.evaluate(body)
            except RefError as e:
                errors.append("%s：{%s} → %s" % (where, m.group(1), e))
                return None
            if not _is_num(v):
                errors.append("%s：{%s} 的值「%s」不是數字" % (where, m.group(1), v))
                return None
            if kinds is not None:
                kinds.append(k)
            return v
        v = to_num(s)
        if v is None:
            errors.append("%s：「%s」不是數字或引用" % (where, s))
        elif kinds is not None:
            kinds.append("percent" if s.endswith("%") else "number")
        return v

    def rows_from(self, ref: str):
        tid, filters, col, agg = parse_ref(ref)
        if col or agg:
            raise RefError("from 只能是表格加篩選，例如 t1[部門=A]")
        t = self.table(tid)
        return t, filter_rows(t, filters)

    def table(self, tid):
        t = self.tables.get(tid)
        if t is None:
            raise RefError("找不到表格 %s（可用：%s）" % (tid, "、".join(self.tables) or "無，請在 spec.sources 填資料檔"))
        return t

    def evaluate(self, body: str):
        body = body.strip()
        if body.startswith("="):
            return self.expr(body[1:])
        if body in self.named:
            return self.named_fact(body)
        if REF_HEAD.match(body):
            return self.ref(body)
        if body.startswith("$"):
            raise RefError("{$key} 只能用在有 repeat 的頁面")
        raise RefError("不是有效的引用（應為 t1…、spec.facts 名稱或 =運算式）")

    # -- internals ------------------------------------------------------
    def ref(self, s: str):
        tid, filters, col, agg = parse_ref(s)
        t = self.table(tid)
        rows = filter_rows(t, filters)
        if col is not None and col not in t.columns:
            raise RefError("%s 沒有欄位「%s」（欄位：%s）" % (tid, col, "、".join(t.columns)))
        if col is None:
            if agg in (None, "count"):
                return len(rows), "number"
            raise RefError("聚合 .%s 需要指定欄位" % agg)
        kind = t.kinds.get(col, "text")
        vals = [r.get(col) for r in rows]
        if agg is None:
            if len(rows) == 1:
                return vals[0], kind
            if not rows:
                raise RefError("篩選後沒有資料列")
            raise RefError("篩選後有 %d 列，請加更多篩選或指定 .sum/.mean 等聚合" % len(rows))
        nonempty = [v for v in vals if v is not None and v != ""]
        if agg == "count":
            return len(nonempty), "number"
        if agg == "distinct":
            return len(set(str(v) for v in nonempty)), "number"
        if agg in ("first", "last"):
            if not nonempty:
                return None, kind
            return (nonempty[0] if agg == "first" else nonempty[-1]), kind
        nums = [v for v in vals if _is_num(v)]
        if not nums:
            raise RefError("欄位「%s」沒有數值可做 .%s" % (col, agg))
        return aggregate(nums, agg), (kind if kind in ("number", "percent") else "number")

    def named_fact(self, name: str):
        if name in self._named_cache:
            return self._named_cache[name]
        if name in self._stack:
            raise RefError("具名數字循環引用：%s" % " → ".join(self._stack + [name]))
        self._stack.append(name)
        try:
            val = self.named[name]
            if _is_num(val):
                res = (val, "number")
            else:
                s = str(val).strip()
                if s.startswith("{") and s.endswith("}"):
                    s = s[1:-1].strip()
                if s.startswith("="):
                    res = self.expr(s[1:])
                elif REF_HEAD.match(s):
                    res = self.ref(s)
                else:
                    n = to_num(s)
                    if n is None:
                        raise RefError("spec.facts.%s 的值「%s」無法解析" % (name, val))
                    res = (n, "percent" if s.endswith("%") else "number")
        finally:
            self._stack.pop()
        self._named_cache[name] = res
        return res

    def expr(self, s: str):
        refs = {}

        def keep(text):
            key = "_r%d" % len(refs)
            refs[key] = text
            return key

        s2 = re.sub(r"`([^`]+)`", lambda m: keep(m.group(1)), s)
        s2 = REF_IN_EXPR.sub(lambda m: keep(m.group(0)), s2)
        try:
            tree = ast.parse(s2.strip(), mode="eval")
        except SyntaxError:
            raise RefError("運算式語法錯誤：%s" % s)
        kinds, flags = [], {"muldiv": False}
        v = self._eval(tree.body, refs, kinds, flags)
        kind = "percent" if kinds and all(k == "percent" for k in kinds) and not flags["muldiv"] else "number"
        return v, kind

    def _eval(self, node, refs, kinds, flags):
        if isinstance(node, ast.Constant):
            if _is_num(node.value):
                return node.value
            raise RefError("運算式只能包含數字")
        if isinstance(node, ast.Name):
            if node.id in refs:
                src = refs[node.id]
                v, k = self.ref(src)
            elif node.id in self.named:
                src = node.id
                v, k = self.named_fact(node.id)
            else:
                raise RefError("運算式中不認得「%s」" % node.id)
            if not _is_num(v):
                raise RefError("「%s」的值「%s」不是數字" % (src, v))
            kinds.append(k)
            return v
        if isinstance(node, ast.BinOp) and type(node.op) in _BIN:
            if isinstance(node.op, (ast.Mult, ast.Div)):
                flags["muldiv"] = True
            a = self._eval(node.left, refs, kinds, flags)
            b = self._eval(node.right, refs, kinds, flags)
            try:
                return _BIN[type(node.op)](a, b)
            except ZeroDivisionError:
                raise RefError("除以 0")
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            v = self._eval(node.operand, refs, kinds, flags)
            return -v if isinstance(node.op, ast.USub) else v
        if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                and node.func.id in _FUNCS and not node.keywords):
            args = [self._eval(a, refs, kinds, flags) for a in node.args]
            return _FUNCS[node.func.id](*args)
        raise RefError("運算式只支援數字、引用、+ - * / 與 abs/round/min/max")
