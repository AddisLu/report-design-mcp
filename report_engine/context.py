"""get_design_context: everything the model needs to write a full spec, in one tool result."""
from __future__ import annotations

import json
from typing import List, Tuple

import yaml

from . import config as C
from .facts import format_value
from .paths import ENGINE_DIR
from .sketch import load_sketch_pages
from .sources import load_sources
from .template import resolve_template_path
from .template_inspector import ensure_catalog

FULL_ROWS = 40
PREVIEW_ROWS = 12
MAX_COLS = 14
DOC_CHARS = 6000


def _guide(name: str) -> str:
    return (ENGINE_DIR / "guides" / name).read_text(encoding="utf-8")


def _cell(v, kind) -> str:
    s = "" if v is None else format_value(v, kind)
    s = s.replace("|", u"／").replace("\n", " ")
    return s[:28] + u"…" if len(s) > 28 else s


KIND_NAMES = {"number": u"數字", "percent": u"百分比", "date": u"日期", "text": u"文字"}


def table_section(t) -> str:
    cols = t.columns[:MAX_COLS]
    lines = [u"### %s ＝ %s（%d 列）" % (t.id, t.label, len(t.rows)),
             u"欄位：" + u"、".join(u"%s(%s)" % (c, KIND_NAMES.get(t.kinds.get(c), u"文字")) for c in t.columns)]
    if len(t.columns) > MAX_COLS:
        lines.append(u"（下表只列前 %d 欄）" % MAX_COLS)
    rows = t.rows if len(t.rows) <= FULL_ROWS else t.rows[:PREVIEW_ROWS]
    lines.append("| " + " | ".join(cols) + " |")
    lines.append("|" + "---|" * len(cols))
    for r in rows:
        lines.append("| " + " | ".join(_cell(r.get(c), t.kinds.get(c)) for c in cols) + " |")
    if len(t.rows) > FULL_ROWS:
        lines.append(u"（只列前 %d 列，共 %d 列；請用引用與聚合取數）" % (PREVIEW_ROWS, len(t.rows)))
    stats, distinct = [], []
    for c in t.columns:
        k = t.kinds.get(c)
        vals = [r.get(c) for r in t.rows if r.get(c) is not None]
        if k in ("number", "percent"):
            nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if nums:
                stats.append(u"%s 合計 %s／平均 %s／最小 %s／最大 %s" % (
                    c, format_value(sum(nums), k), format_value(sum(nums) / float(len(nums)), k),
                    format_value(min(nums), k), format_value(max(nums), k)))
        elif k == "text":
            uniq = list(dict.fromkeys(str(v) for v in vals))
            if 1 < len(uniq) <= 15 and len(uniq) < len(t.rows):
                distinct.append(u"%s＝%s" % (c, u"、".join(uniq)))
    if stats:
        lines.append(u"數值摘要：" + u"；".join(stats))
    if distinct:
        lines.append(u"可篩選值：" + u"；".join(distinct))
    return "\n".join(lines)


def layouts_section(cat: dict) -> str:
    lines = [u"## 範本版型（%s，%s × %s 吋）" % (cat.get("template", ""), *cat.get("slide_size_in", ["?", "?"])),
             u"| layout key | 用途 | 內容區（格線 x,y,w,h／約可放字數） |", "|---|---|---|"]
    for L in cat.get("layouts", []):
        areas = u"；".join(u"%s %s／%d字" % (c["kind"], ",".join(str(v) for v in c["grid"]), c["capacity_chars"])
                          for c in L.get("content", [])) or u"—"
        lines.append(u"| %s | %s | %s |" % (L["key"], L.get("use", ""), areas))
    lines.append(u"自動挑選：封面＝%s；1 區＝%s；2 區＝%s；自訂排版底版＝%s" % (
        cat.get("cover"), cat.get("by_count", {}).get("1"), cat.get("by_count", {}).get("2"),
        cat.get("custom_base", {}).get("with_title")))
    return "\n".join(lines)


def blueprints_section(bps: dict, chosen=None) -> str:
    if not bps:
        return u"## 已存版型\n（還沒有。使用者滿意某份草稿版型時，可用 save_layout 存起來重複使用。）"
    lines = [u"## 已存版型"]
    for name, bp in bps.items():
        lines.append(u"- **%s**（%d 頁）：%s" % (name, len(bp.get("slides") or []), bp.get("description", "")))
    if chosen:
        if chosen not in bps:
            lines.append(u"⚠ 找不到已存版型「%s」" % chosen)
        else:
            lines.append(u"\n### 「%s」完整骨架（每頁用 layout:\"%s/頁碼\" ＋ fill 依 region id 填內容；directive 是該區塊的預設指示）"
                         % (chosen, chosen))
            lines.append("```json\n%s\n```" % json.dumps(bps[chosen], ensure_ascii=False, indent=1))
    return "\n".join(lines)


def build_context(audience=None, sources=None, sketch=None, sketch_pages=None, blueprint=None,
                  template=None) -> Tuple[str, List[dict]]:
    cfg = C.load_config()
    aud_key, aud = C.get_audience(audience)
    tpath = resolve_template_path(template, cfg)
    cat = ensure_catalog(tpath)
    bundle = load_sources(sources or [])
    pages = load_sketch_pages(sketch, sketch_pages) if sketch else []

    src_json = json.dumps(list(sources or []), ensure_ascii=False)
    parts = [
        u"# 報告設計情境（ReportDesign）",
        u"**下一步**：讀完以下內容後，直接寫出完整 spec 並呼叫 `build_deck`"
        u"（spec.audience=\"%s\"，spec.sources=%s%s），再用 `preview_deck` 看圖自我檢查。"
        u"不要先停下來請使用者確認大綱。"
        % (aud_key, src_json, u"，spec.template=\"%s\"" % template if template else ""),
    ]
    others = [k for k in C.list_audiences() if k != aud_key]
    parts.append(u"## 受眾：%s（其他可用：%s）\n```yaml\n%s```" % (
        aud_key, u"、".join(others) or u"無", yaml.safe_dump(aud, allow_unicode=True, sort_keys=False)))
    guide = C.style_guide_text()
    if guide:
        parts.append(u"## 文風指南\n" + guide.strip())
    parts.append(layouts_section(cat))
    parts.append(blueprints_section(C.load_blueprints(), blueprint))

    data = [u"## 資料"]
    if not bundle.tables and not bundle.documents:
        data.append(u"（沒有提供資料來源。若使用者有資料檔，請把路徑放進 sources 重新呼叫本工具。）")
    for t in bundle.tables.values():
        data.append(table_section(t))
    for d in bundle.documents:
        text = d.text.strip()
        more = u"\n…（共 %d 字，只列前 %d 字）" % (len(text), DOC_CHARS) if len(text) > DOC_CHARS else ""
        data.append(u"### %s ＝ %s（文字，不能用 {} 引用；需要的數字請寫進 spec.facts）\n```\n%s%s\n```"
                    % (d.id, d.source, text[:DOC_CHARS], more))
    for e in bundle.errors:
        data.append(u"⚠ " + e)
    parts.append("\n\n".join(data))

    parts.append(_guide("visual_guide.md"))
    parts.append(_guide("spec_guide.md"))
    exs = C.examples()
    if exs:
        chosen = next((e for e in exs if aud_key in e[0].stem), exs[0])
        parts.append(u"## 範例 spec（%s，示範寫法；實際內容請依本次資料）\n```json\n%s\n```"
                     % (chosen[0].name, json.dumps(chosen[1], ensure_ascii=False, indent=1)))
    if pages:
        parts.append(u"## 手繪草稿（共 %d 頁，下方圖片依序為：%s）" % (
            len(pages), u"、".join(u"第 %d 頁" % p["page"] for p in pages)))
        parts.append(_guide("sketch_guide.md"))
    return "\n\n".join(parts), pages
