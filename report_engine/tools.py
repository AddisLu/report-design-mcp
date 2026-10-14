"""MCP tool definitions and handlers (transport-independent)."""
from __future__ import annotations

import datetime as _dt
import json
import sys
import traceback

from . import config as C
from . import deck
from .context import build_context
from . import preview
from .deck import REGION_TYPES
from .lint import format_lint, lint_deck
from .paths import ENGINE_DIR, resolve_path
from .sketch import load_sketch_pages, make_canvas
from .template_inspector import ensure_catalog

SERVER_INSTRUCTIONS = (
    u"製作 PowerPoint 報告用。流程：1) get_design_context（audience、sources；有手繪草稿就帶 sketch）；"
    u"2) 讀完直接寫出完整 spec 呼叫 build_deck；3) 呼叫 preview_deck 看每頁圖片與版面檢查，"
    u"用 patch_deck 只修有問題的頁（自我檢查最多 2 輪）；4) 回報使用者。數字一律用 {t1…} 引用。"
)

_STR = {"type": "string"}

TOOL_DEFS = [
    {
        "name": "get_design_context",
        "description": (
            u"做報告的第一步（通常也是唯一一次查詢）。一次回傳：受眾規則、文風指南、範本版型、已存版型、"
            u"資料摘要與表格 id（t1…）、spec 寫法、範例；有 sketch 時附上所有草稿頁圖片與判讀規則。"
            u"讀完就直接呼叫 build_deck。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "audience": {"type": "string", "description": u"executive / manager / engineer，或「主管」「經理」「工程師」"},
                "sources": {"type": "array", "items": _STR,
                            "description": u"資料檔或資料夾：xlsx、csv、docx、pptx、pdf、md、txt"},
                "sketch": {"type": "string", "description": u"手繪草稿：多頁 PDF、圖片，或放圖片的資料夾"},
                "sketch_pages": {"type": "string", "description": u"只讀部分頁，例如 1-8（頁數很多時分批）"},
                "blueprint": {"type": "string", "description": u"已存版型名稱（使用者說「用 XX 版型」時）"},
                "template": {"type": "string", "description": u"指定 .pptx/.potx 範本；省略用預設"},
            },
        },
    },
    {
        "name": "build_deck",
        "description": (
            u"依 spec 一次完成檢查、自動修正與產出 .pptx（並另存 .spec.json 供 patch_deck 使用）。"
            u"只有引用錯誤這類無法自動修的問題才會退回，照列出的地方修正後重送一次即可。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "spec": {"type": "object", "description": u"完整 spec（寫法見 get_design_context）", "additionalProperties": True},
                "out_path": {"type": "string", "description": u"輸出 .pptx 路徑或資料夾；省略存到 output/"},
            },
            "required": ["spec"],
        },
    },
    {
        "name": "patch_deck",
        "description": u"只修改指定頁再重新產出（其他頁不動）。頁碼是 spec 的頁碼（build_deck 回傳的對照表左欄）。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "spec_path": {"type": "string", "description": u"build_deck 回傳的 .spec.json"},
                "slides": {"type": "object", "additionalProperties": True,
                           "description": u"{\"3\": {要改的欄位…}}；值為 null 表示刪除該欄位；加 \"_replace\": true 表示整頁取代"},
                "delete": {"type": "array", "items": {"type": "integer"}, "description": u"要刪除的頁碼"},
                "insert": {"type": "array", "items": {"type": "object"},
                           "description": u"[{\"after\": 2, \"slide\": {…}}]，after=0 表示放最前面"},
                "out_path": {"type": "string"},
            },
            "required": ["spec_path"],
        },
    },
    {
        "name": "preview_deck",
        "description": (
            u"build_deck／patch_deck 之後的自我檢查：回傳版面檢查（溢出、重疊、對比、字級不一致、太密、太空）"
            u"與每頁投影片圖片（用本機 PowerPoint 匯出）。看圖找出問題後用 patch_deck 只改那幾頁。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": u"build_deck 回傳的 .pptx（或 .spec.json）路徑"},
                "slides": {"type": "string", "description": u"只看部分投影片，例如 1-6 或 2,5；一次最多 12 頁"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "view_image",
        "description": (
            u"查看要放進簡報的截圖或照片（可給檔案或資料夾）。圖片上疊了 0.1 間隔的格線與刻度，"
            u"用來決定 image 的 highlights／zoom／crop 座標（0–1 比例）。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string", "description": u"圖片檔或資料夾（最多 8 張）"}},
            "required": ["path"],
        },
    },
    {
        "name": "read_sketch",
        "description": u"讀取手繪草稿頁面圖片（多頁 PDF、圖片或資料夾）。通常直接用 get_design_context 的 sketch 參數即可；頁數很多時用這個分批讀。",
        "inputSchema": {
            "type": "object",
            "properties": {"path": _STR, "pages": {"type": "string", "description": u"例如 9-16"}},
            "required": ["path"],
        },
    },
    {
        "name": "get_sketch_canvas",
        "description": u"產生手寫板用的 12×8 格線底稿（.pdf 多頁或 .png），可選擇第一頁加「需求」頁。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "out_path": {"type": "string", "description": u"例如 sketches/canvas.pdf"},
                "pages": {"type": "integer", "description": u"版面頁數，預設 6"},
                "requirements_page": {"type": "boolean", "description": u"第一頁加需求頁，預設 true"},
            },
            "required": ["out_path"],
        },
    },
    {
        "name": "save_layout",
        "description": (
            u"把滿意的草稿版型存成可重複使用的已存版型（只存框線、類型與「每次都適用」的指示，不存這次的內容或一次性註解）。"
            u"之後使用者說「用 XX 版型」就不必再畫。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": u"版型名稱，例如「月報-主管版」"},
                "description": _STR,
                "audience": _STR,
                "slides": {
                    "type": "array",
                    "description": u"每頁：{\"regions\":[{\"id\":\"r1\",\"type\":\"kpi\",\"x\":0,\"y\":1.5,\"w\":12,\"h\":2,\"directive\":\"…\"}],"
                                   u"\"directive\":\"…\",\"repeat\":{…}}",
                    "items": {"type": "object"},
                },
            },
            "required": ["name", "slides"],
        },
    },
    {
        "name": "inspect_template",
        "description": u"解析新的公司範本（.pptx/.potx），更新版型表；set_default=true 則設為預設範本。",
        "inputSchema": {
            "type": "object",
            "properties": {"template_path": _STR, "set_default": {"type": "boolean"}},
            "required": ["template_path"],
        },
    },
]


def _text(s, is_error=False):
    return {"content": [{"type": "text", "text": s}], "isError": is_error}


def _obj(v):
    if isinstance(v, str):
        return json.loads(v)
    return v


def _images(pages):
    return [{"type": "image", "data": p["data"], "mimeType": "image/jpeg"} for p in pages]


def format_build(res) -> str:
    if not res.get("ok"):
        lines = [u"❌ 尚未產出。請只修正以下 %d 個問題，再重送一次 build_deck：" % len(res["errors"])]
        lines += [u"- " + e for e in res["errors"]]
        if res.get("warnings"):
            lines.append(u"\n其他提醒：")
            lines += [u"- " + w for w in res["warnings"][:15]]
        return "\n".join(lines)
    lines = [u"✅ 已產出：%s（%d 張投影片）" % (res["path"], res["slide_count"]),
             u"spec：%s（之後用 patch_deck 修改）" % res["spec_path"],
             u"範本：%s；受眾：%s" % (res["template"], res["audience"]),
             "", u"### 頁面對照（spec 頁 → 投影片）"]
    for m in res["mapping"]:
        nums = m["slides"]
        rng = u"—" if not nums else (str(nums[0]) if len(nums) == 1 else u"%d–%d" % (nums[0], nums[-1]))
        extra = u"（展開：%s）" % u"、".join(m["repeat"]) if m["repeat"] else ""
        cont = u"（含續頁）" if len(nums) > 1 and not m["repeat"] else ""
        lines.append(u"- spec %d → %s%s%s" % (m["spec"], rng, extra, cont))
    if res["annotations"]:
        lines += ["", u"### 註解解讀表", u"| 草稿頁 | 投影片 | 手寫辨識 | 套用方式 |", "|---|---|---|---|"]
        for a in res["annotations"]:
            sl = u"、".join(str(n) for n in a["slides"]) or u"—"
            lines.append(u"| %s | %s | %s | %s |" % (a["page"], sl, a["raw"], a["applied"]))
    if res["warnings"]:
        lines += ["", u"### 提醒（%d）" % len(res["warnings"])]
        lines += [u"- " + w for w in res["warnings"][:25]]
        if len(res["warnings"]) > 25:
            lines.append(u"- …另有 %d 則" % (len(res["warnings"]) - 25))
    if res.get("lint"):
        lines += ["", u"### " + format_lint(res["lint"])]
    lines += ["", u"下一步：呼叫 preview_deck(path=\"%s\") 看每頁圖片自我檢查（對照 get_design_context 的視覺設計指引），"
                  u"只針對 ⚠ 項目與圖上看得到的問題用 patch_deck 修正，最多 2 輪；"
                  u"之後把檔案路徑、註解解讀表與重要提醒簡短告訴使用者。" % res["path"]]
    return "\n".join(lines)


def h_get_design_context(a):
    text, pages = build_context(a.get("audience"), _obj(a.get("sources")) or [], a.get("sketch"),
                                a.get("sketch_pages"), a.get("blueprint"), a.get("template"))
    return {"content": [{"type": "text", "text": text}] + _images(pages), "isError": False}


def h_build_deck(a):
    res = deck.build(_obj(a.get("spec")), a.get("out_path"))
    return _text(format_build(res), not res.get("ok"))


def h_patch_deck(a):
    res = deck.patch(a["spec_path"], _obj(a.get("slides")), _obj(a.get("delete")), _obj(a.get("insert")), a.get("out_path"))
    return _text(format_build(res), not res.get("ok"))


def h_preview_deck(a):
    p = resolve_path(a["path"])
    if p.suffix.lower() == ".json":
        out = (C.read_json(p, {}) or {}).get("_output")
        p = resolve_path(out) if out else p.with_suffix(".pptx")
    if not p.exists():
        return _text(u"找不到簡報：%s" % a["path"], True)
    res = lint_deck(p)
    text = format_lint(res)
    imgs, err = preview.export_images(p, a.get("slides"), res["slide_count"])
    if imgs:
        text += u"\n\n下方圖片依序為：%s。檢查：每頁一個重點是否一眼看懂、層級（標題＞重點數字＞說明）、" \
                u"對齊與留白、顏色是否只用在重點、有沒有文字被截斷或擠在一起。" % u"、".join(u"投影片 %d" % i["slide"] for i in imgs)
    else:
        text += u"\n\n（沒有預覽圖：%s。請依上面的版面檢查修正。）" % (err or u"未知原因")
    content = [{"type": "text", "text": text}]
    content += [{"type": "image", "data": i["data"], "mimeType": "image/jpeg"} for i in imgs or []]
    return {"content": content, "isError": False}


def h_view_image(a):
    from .images import grid_preview

    p = resolve_path(a["path"])
    exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff")
    files = sorted(x for x in p.iterdir() if x.suffix.lower() in exts) if p.is_dir() else [p]
    files = [f for f in files if f.exists()][:8]
    if not files:
        return _text(u"找不到圖片：%s" % a["path"], True)
    lines, imgs = [u"下方圖片依序如下（紅線是 0.1 間隔格線，座標以原圖比例 0–1 表示，左上角是 0,0）："], []
    for f in files:
        g = grid_preview(f)
        lines.append(u"- %s：%d×%d px" % (f, g["size"][0], g["size"][1]))
        imgs.append({"type": "image", "data": g["data"], "mimeType": "image/jpeg"})
    lines.append(u"用法：image 的 highlights 寫 {\"x\":0.14,\"y\":0.27,\"w\":0.8,\"h\":0.06,\"note\":\"…\"}；"
                 u"系統截圖建議 style:\"browser\"。")
    return {"content": [{"type": "text", "text": "\n".join(lines)}] + imgs, "isError": False}


def h_read_sketch(a):
    pages = load_sketch_pages(a["path"], a.get("pages"))
    guide = (ENGINE_DIR / "guides" / "sketch_guide.md").read_text(encoding="utf-8")
    head = u"草稿 %s：本次 %d 頁（%s）\n\n%s" % (
        a["path"], len(pages), u"、".join(u"第 %d 頁" % p["page"] for p in pages), guide)
    return {"content": [{"type": "text", "text": head}] + _images(pages), "isError": False}


def h_get_sketch_canvas(a):
    paths = make_canvas(a["out_path"], int(a.get("pages") or 6), a.get("requirements_page", True) is not False)
    return _text(u"已產生底稿：\n" + "\n".join(u"- " + p for p in paths) +
                 u"\n匯入手寫板 App 當底圖：黑筆畫框並寫類型，紅筆寫註解，左上角標頁面類型，右上角寫頁碼。")


def h_save_layout(a):
    name = str(a.get("name", "")).strip()
    if not name or "/" in name:
        return _text(u"版型名稱不可空白，也不能含「/」", True)
    slides, errors = [], []
    for i, s in enumerate(_obj(a.get("slides")) or [], 1):
        regs = []
        for j, r in enumerate(s.get("regions") or [], 1):
            t = str(r.get("type", "")).strip()
            if not (REGION_TYPES.get(t.lower()) or REGION_TYPES.get(t)):
                errors.append(u"第 %d 頁區塊 %d：不認得類型「%s」" % (i, j, t))
                continue
            try:
                box = {k: round(float(r[k]), 2) for k in ("x", "y", "w", "h")}
            except (KeyError, TypeError, ValueError):
                errors.append(u"第 %d 頁區塊 %d：缺少 x,y,w,h" % (i, j))
                continue
            reg = {"id": str(r.get("id") or "r%d" % j), "type": t}
            reg.update(box)
            for k in ("directive", "label"):
                if r.get(k):
                    reg[k] = r[k]
            regs.append(reg)
        page = {"regions": regs}
        for k in ("directive", "repeat", "page_type", "title_hint"):
            if s.get(k):
                page[k] = s[k]
        slides.append(page)
    if errors:
        return _text(u"沒有存檔：\n" + "\n".join(u"- " + e for e in errors), True)
    bps = C.load_blueprints()
    existed = name in bps
    bps[name] = {"description": a.get("description", ""), "audience": a.get("audience", ""),
                 "slides": slides, "saved_at": _dt.datetime.now().isoformat(timespec="seconds")}
    C.save_blueprints(bps)
    return _text(u"已%s版型「%s」（%d 頁）。之後可用 get_design_context(blueprint=\"%s\")，"
                 u"spec 每頁寫 layout:\"%s/頁碼\" 加 fill:{region id: 內容}。"
                 % (u"更新" if existed else u"儲存", name, len(slides), name, name))


def h_inspect_template(a):
    p = resolve_path(a["template_path"])
    if not p.exists():
        return _text(u"找不到範本：%s" % a["template_path"], True)
    cat = ensure_catalog(p, force=True)
    if a.get("set_default"):
        cfg = C.load_config()
        cfg["template"] = str(p)
        C.save_config(cfg)
    lines = [u"範本 %s：%d 個版型%s" % (p.name, len(cat["layouts"]), u"（已設為預設）" if a.get("set_default") else "")]
    for L in cat["layouts"]:
        lines.append(u"- %s：%s" % (L["key"], L["use"]))
    lines.append(u"版型用途說明可在 design_system/layout_catalog.json 的 use 欄位手動修改。")
    return _text("\n".join(lines))


HANDLERS = {
    "get_design_context": h_get_design_context,
    "build_deck": h_build_deck,
    "patch_deck": h_patch_deck,
    "preview_deck": h_preview_deck,
    "view_image": h_view_image,
    "read_sketch": h_read_sketch,
    "get_sketch_canvas": h_get_sketch_canvas,
    "save_layout": h_save_layout,
    "inspect_template": h_inspect_template,
}


def call_tool(name, args) -> dict:
    fn = HANDLERS.get(name)
    if fn is None:
        return _text(u"沒有這個工具：%s" % name, True)
    try:
        return fn(args or {})
    except Exception as e:  # report to the model instead of failing the JSON-RPC call
        traceback.print_exc(file=sys.stderr)
        return _text(u"%s 失敗：%s: %s" % (name, type(e).__name__, e), True)
