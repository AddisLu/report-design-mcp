"""Designed diagram blocks drawn with native shapes: gantt (甘特圖), timeline (時程表), process (流程)."""
from __future__ import annotations

import datetime as dt
import math

from pptx.enum.shapes import MSO_SHAPE

from .canvas import EMU_PT, Canvas
from .theme import mix

EMU_IN = 914400
DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y%m%d", "%Y-%m", "%Y/%m")


def to_date(v):
    """date / datetime / '2026/10/14' / '2026-10' → date; raises ValueError."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(u"看不懂的日期「%s」（請用 2026/10/14 或 2026-10-14）" % s)


def _month_start(d):
    return dt.date(d.year, d.month, 1)


def _add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    return dt.date(d.year + y, m + 1, 1)


def _em(text):
    from .renderer import em_width, plain_text
    return em_width(plain_text(text or ""))


class _Scale:
    """Date → x (EMU) over [lo, hi)."""

    def __init__(self, lo, hi, x0, x1):
        self.lo, self.hi, self.x0, self.x1 = lo, hi, x0, x1
        self.days = max(1, (hi - lo).days)

    def x(self, d):
        return self.x0 + (self.x1 - self.x0) * ((d - self.lo).days / float(self.days))

    def ticks(self):
        """[(date, label, major)] period starts + the unit name."""
        span = self.days
        out = []
        if span <= 62:
            d = self.lo - dt.timedelta(days=self.lo.weekday())
            while d < self.hi:
                if d >= self.lo:
                    out.append((d, "%d/%d" % (d.month, d.day), d.day <= 7))
                d += dt.timedelta(days=7)
            return out, "week"
        if span > 550:
            d = _month_start(self.lo)
            while (d.month - 1) % 3:
                d = _add_months(d, -1)
            while d < self.hi:
                if d >= self.lo:
                    q = (d.month - 1) // 3 + 1
                    out.append((d, ("%d Q%d" % (d.year, q)) if q == 1 or not out else "Q%d" % q, q == 1))
                d = _add_months(d, 3)
            return out, "quarter"
        d = _month_start(self.lo)
        while d < self.hi:
            first = not out or d.month == 1
            out.append((d, ("%d/%d" % (d.year, d.month)) if first else (u"%d月" % d.month), first))
            d = _add_months(d, 1)
        return out, "month"


def _status_color(theme, key, default):
    return theme.status(key) if key in ("warn", "bad") else default


def _title(r, slide, rect, b):
    return r._chart_title(slide, rect, b["title"]) if b.get("title") else rect


# -- gantt -------------------------------------------------------------------
def draw_gantt(r, slide, rect, b):
    th = r.theme
    x, y, w, h = _title(r, slide, rect, b)
    tasks = list(b["tasks"])
    rows = []
    groups = []
    for t in tasks:
        g = t.get("group") or ""
        if g and g not in groups:
            groups.append(g)
    if groups:
        for g in groups:
            members = [t for t in tasks if (t.get("group") or "") == g]
            rows.append({"kind": "group", "name": g, "start": min(t["start"] for t in members),
                         "end": max(t["end"] for t in members)})
            rows.extend(dict(t, kind="task") for t in members)
        rows.extend(dict(t, kind="task") for t in tasks if not t.get("group"))
    else:
        rows = [dict(t, kind="task") for t in tasks]
    for m in b.get("milestones") or []:
        rows.append({"kind": "task", "name": m["name"], "start": m["date"], "end": m["date"], "milestone": True,
                     "status": m.get("status")})

    header_h = EMU_PT * 26
    avail = h - header_h - EMU_PT * 22  # room for the "today" tag under the chart
    row_h = min(EMU_IN * 0.5, avail / float(max(1, len(rows))))
    if row_h < EMU_IN * 0.2:
        keep = max(1, int(avail // (EMU_IN * 0.2)))
        r.warnings.append(u"%s：甘特圖 %d 列放不下，只顯示前 %d 列（可拆頁或用 group 合併）" % (r._where, len(rows), keep))
        rows = rows[:keep]
        row_h = avail / float(keep)
    rh_pt = row_h / EMU_PT
    font = th.snap(max(10.0, min(th.size("body", 2), rh_pt * 0.5)))
    small = th.snap(max(9.0, font - 2))
    has_owner = any(rw.get("owner") for rw in rows)
    name_w = max(_em(rw["name"]) * font for rw in rows) * EMU_PT + EMU_PT * (14 if groups else 6)
    owner_w = (max(_em(rw.get("owner")) for rw in rows) * small * EMU_PT + EMU_PT * 12) if has_owner else 0
    label_w = min(w * 0.34, name_w + owner_w + EMU_PT * 10)
    gx0, gx1 = x + label_w, x + w - EMU_PT * 30  # room for the progress label at the end
    lo = min(rw["start"] for rw in rows)
    hi = max(rw["end"] for rw in rows) + dt.timedelta(days=1)
    today = b.get("today")
    sc = _Scale(lo, hi, gx0, gx1)
    ticks, unit = sc.ticks()
    if unit == "month":
        sc = _Scale(_month_start(lo), _add_months(_month_start(hi - dt.timedelta(days=1)), 1), gx0, gx1)
        ticks, unit = sc.ticks()

    cv = Canvas(r, slide, "Gantt")
    top = y + header_h
    bottom = top + row_h * len(rows)
    # zebra rows and group bands
    for i, rw in enumerate(rows):
        ry = top + i * row_h
        if rw["kind"] == "group":
            cv.rect(x, ry, w, row_h, fill=th.neutral(100))
        elif i % 2 == 1:
            cv.rect(x, ry, w, row_h, fill=th.neutral(50))
    # period grid + header labels
    for i, (d, label, major) in enumerate(ticks):
        tx = sc.x(d)
        nxt = sc.x(ticks[i + 1][0]) if i + 1 < len(ticks) else sc.x1
        cv.line(tx, top, tx, bottom, th.neutral(300) if major else th.neutral(100), 0.75)
        cv.text((tx + nxt) / 2.0, y + header_h / 2.0, label, small, color=th.neutral(700) if major else th.neutral(500),
                bold=major, align="middle", valign="middle")
    cv.line(x, top, x + w, top, th.neutral(300), 1.0)
    cv.line(x, bottom, x + w, bottom, th.neutral(300), 0.75)
    # rows
    for i, rw in enumerate(rows):
        ry = top + i * row_h
        cy = ry + row_h / 2.0
        indent = EMU_PT * 10 if groups and rw["kind"] == "task" else EMU_PT * 4
        cv.text(x + indent, cy, rw["name"], font, color=th.neutral(900), bold=rw["kind"] == "group", valign="middle",
                width=name_w - indent)
        if has_owner and rw.get("owner"):
            cv.text(x + label_w - EMU_PT * 8, cy, rw["owner"], small, color=th.neutral(500), align="end", valign="middle")
        x0, x1 = sc.x(rw["start"]), sc.x(rw["end"] + dt.timedelta(days=1))
        if rw["kind"] == "group":  # summary bracket
            bh = row_h * 0.16
            cv.rect(x0, cy - bh / 2.0, x1 - x0, bh, fill=th.neutral(700))
            for ex in (x0, x1 - bh):
                cv.rect(ex, cy - bh / 2.0, bh, row_h * 0.32, fill=th.neutral(700))
            continue
        color = _status_color(th, rw.get("status"), th.primary)
        if rw.get("milestone") or rw["start"] == rw["end"] and rw.get("progress") is None:
            s = row_h * 0.46
            mx = sc.x(rw["start"])
            cv.autoshape(MSO_SHAPE.DIAMOND, mx - s / 2.0, cy - s / 2.0, s, s, fill=color)
            cv.text(mx + s * 0.75, cy, u"%d/%d" % (rw["start"].month, rw["start"].day), small, color=th.neutral(700),
                    valign="middle")
            continue
        bh = row_h * 0.52
        rad = bh * 0.3
        prog = rw.get("progress")
        if prog is None:
            cv.rect(x0, cy - bh / 2.0, x1 - x0, bh, fill=color, radius=rad)
        else:
            cv.rect(x0, cy - bh / 2.0, x1 - x0, bh, fill=mix(color, "FFFFFF", 0.75), radius=rad)
            pw = (x1 - x0) * max(0.0, min(1.0, prog))
            if pw > rad * 2:
                cv.rect(x0, cy - bh / 2.0, pw, bh, fill=color, radius=rad)
            label_color = th.status_text(rw["status"]) if rw.get("status") in ("warn", "bad") else th.neutral(700)
            cv.text(x1 + EMU_PT * 4, cy, "%d%%" % round(prog * 100), small, color=label_color, valign="middle",
                    bold=rw.get("status") == "bad")
    if today and lo <= today <= hi:
        tx = sc.x(today)
        cv.line(tx, top, tx, bottom, th.alert, 1.25, dash=True)
        tag = u"今天 %d/%d" % (today.month, today.day)
        lw = (_em(tag) * small + 12) * EMU_PT
        pill = cv.rect(tx - lw / 2.0, bottom + EMU_PT * 4, lw, small * EMU_PT * 1.5, fill=th.alert,
                       radius=small * EMU_PT * 0.75)
        cv.text(0, 0, tag, small, color="FFFFFF", bold=True, align="middle", shape=pill)


# -- timeline ----------------------------------------------------------------
def draw_timeline(r, slide, rect, b):
    th = r.theme
    x, y, w, h = _title(r, slide, rect, b)
    items = list(b["items"])
    n = len(items)
    if not n:
        return
    pad = EMU_PT * 12
    x0, x1 = x + pad + EMU_PT * 40, x + w - pad - EMU_PT * 40
    alternate = n >= 5
    axis_y = y + h / 2.0 if alternate else y + h * 0.28
    dated = all(it.get("date") for it in items)
    if dated and not b.get("even"):
        lo, hi = min(it["date"] for it in items), max(it["date"] for it in items)
        span = max(1, (hi - lo).days)
        xs = [x0 + (x1 - x0) * (it["date"] - lo).days / float(span) for it in items] if n > 1 else [(x0 + x1) / 2.0]
        min_gap = (x1 - x0) / float(n) * (0.55 if alternate else 0.85)
        if any(b2 - a < min_gap for a, b2 in zip(xs, xs[1:])):
            xs = None
    else:
        xs = None
    if xs is None:
        xs = [x0 + (x1 - x0) * i / float(max(1, n - 1)) for i in range(n)] if n > 1 else [(x0 + x1) / 2.0]
    slot = (x1 - x0) / float(max(1, n - 1)) if n > 1 else (x1 - x0)
    block_w = min(slot * (1.85 if alternate else 0.95), EMU_IN * 2.6)
    cv = Canvas(r, slide, "Timeline")
    cv.line(x + pad, axis_y, x + w - pad, axis_y, th.neutral(300), 2.0)
    today = b.get("today")
    done_x = None
    if today and dated:
        before = [xi for xi, it in zip(xs, items) if it["date"] <= today]
        if before:
            done_x = max(before)
    if done_x is None:
        done_idx = [i for i, it in enumerate(items) if it.get("state") == "done"]
        if done_idx:
            done_x = xs[max(done_idx)]
    if done_x is not None:
        cv.line(x + pad, axis_y, done_x, axis_y, th.primary, 3.0)
    head = th.snap(th.size("body", 1))
    small = th.snap(max(11.0, head - 3))
    date_size = th.snap(max(11.0, head - 2))
    stem = EMU_PT * 18
    lh = 1.3 * EMU_PT
    for i, (cx, it) in enumerate(zip(xs, items)):
        state = it.get("state") or ("done" if today and it.get("date") and it["date"] <= today else "todo")
        status = it.get("status")
        dot = EMU_PT * 7
        if status in ("warn", "bad"):
            cv.ellipse(cx, axis_y, dot, dot, fill=th.status(status), stroke="FFFFFF", width=2)
        elif state == "done":
            cv.ellipse(cx, axis_y, dot, dot, fill=th.primary, stroke="FFFFFF", width=2)
        elif state == "current":
            cv.ellipse(cx, axis_y, dot * 1.6, dot * 1.6, fill=mix(th.primary, "FFFFFF", 0.78))
            cv.ellipse(cx, axis_y, dot, dot, fill=th.primary, stroke="FFFFFF", width=2)
        else:
            cv.ellipse(cx, axis_y, dot, dot, fill="FFFFFF", stroke=th.neutral(300), width=2)
        up = alternate and i % 2 == 0
        bx = min(max(cx, x + block_w / 2.0), x + w - block_w / 2.0)  # keep the text block on the slide
        when = it.get("when") or (u"%d/%d" % (it["date"].month, it["date"].day) if it.get("date") else "")
        label_lines = max(1, int(math.ceil(_em(it["label"]) * head * EMU_PT * 1.1 / max(block_w, 1))))
        note_lines = max(1, int(math.ceil(_em(it.get("note")) * small * EMU_PT * 1.15 / max(block_w, 1)))) if it.get("note") else 0
        parts = []
        if when:
            accent = th.status_text(status) if status in ("warn", "bad") else (th.primary if state != "todo" else th.neutral(500))
            parts.append((when, date_size, accent, True, 1))
        parts.append((it["label"], head, th.neutral(900), True, label_lines))
        if note_lines:
            parts.append((it["note"], small, th.neutral(700), False, note_lines))
        if up:
            cv.line(cx, axis_y - dot, cx, axis_y - stem + EMU_PT * 2, th.neutral(300), 0.75)
            ty = axis_y - stem
            for text, size, color, bold, n_lines in reversed(parts):
                bh = n_lines * size * lh + EMU_PT * 2
                cv.text(bx, ty, text, size, color=color, bold=bold, align="middle", valign="bottom", width=block_w,
                        height=bh, wrap=True)
                ty -= bh
        else:
            cv.line(cx, axis_y + dot, cx, axis_y + stem - EMU_PT * 2, th.neutral(300), 0.75)
            ty = axis_y + stem
            for text, size, color, bold, n_lines in parts:
                bh = n_lines * size * lh + EMU_PT * 2
                cv.text(bx, ty, text, size, color=color, bold=bold, align="middle", width=block_w, height=bh, wrap=True)
                ty += bh


# -- process -----------------------------------------------------------------
def draw_process(r, slide, rect, b):
    th = r.theme
    x, y, w, h = _title(r, slide, rect, b)
    steps = list(b["steps"])
    n = len(steps)
    if not n:
        return
    has_notes = any(s.get("note") or s.get("owner") for s in steps)
    num_size = th.snap(th.size("caption") + 1)
    chev_h = min(EMU_IN * 0.85, h * (0.32 if has_notes else 0.6))
    gap = EMU_PT * 3
    point = chev_h * 0.32
    cw = (w + point * (n - 1) - gap * (n - 1)) / float(n)
    top = y + num_size * EMU_PT * 1.8 if has_notes else y + (h - chev_h) / 2.0
    states = [s.get("state") for s in steps]
    explicit = any(states)
    cv = Canvas(r, slide, "Process")
    title_size = th.size("body", 1)
    for s in steps:
        title_size = min(title_size, (cw - point * 2 - EMU_PT * 10) / EMU_PT / max(_em(s["title"]), 1.0))
    title_size = th.snap(max(11.0, title_size))
    note_size = th.snap(max(10.0, th.size("body", 2) - (2 if n > 4 else 0)))
    for i, s in enumerate(steps):
        sx = x + i * (cw - point + gap)
        state, status = s.get("state"), s.get("status")
        if status in ("warn", "bad"):
            fill, fg = th.status(status), "FFFFFF"
        elif not explicit:
            fill, fg = mix(th.primary, "FFFFFF", 0.5 * i / float(max(1, n - 1))), "FFFFFF"
        elif state == "done":
            fill, fg = th.primary, "FFFFFF"
        elif state == "current":
            fill, fg = th.secondary, "FFFFFF"
        else:
            fill, fg = th.neutral(100), th.neutral(700)
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        shp = cv.autoshape(kind, sx, top, cw, chev_h, fill=fill, adjust=point / float(chev_h))
        cv.text(0, 0, s["title"], title_size, color=fg, bold=True, align="middle", shape=shp)
        if has_notes:
            label = "%02d" % (i + 1)
            cv.text(sx + EMU_PT * 2, y, label, num_size, color=th.primary if state != "todo" else th.neutral(500),
                    bold=True)
            ny = top + chev_h + EMU_PT * 10
            nw = cw - point - EMU_PT * 4
            if s.get("note"):
                lines = max(1, int(math.ceil(_em(s["note"]) * note_size * EMU_PT / max(nw, 1))))
                cv.text(sx + EMU_PT * 2, ny, s["note"], note_size, color=th.neutral(700), width=nw, wrap=True,
                        height=lines * note_size * 1.35 * EMU_PT + EMU_PT * 2)
                ny += lines * note_size * 1.35 * EMU_PT + EMU_PT * 6
            if s.get("owner"):
                cv.text(sx + EMU_PT * 2, ny, s["owner"], th.snap(note_size - 1), color=th.neutral(500), width=nw)
