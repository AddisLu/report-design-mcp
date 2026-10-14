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
    u"2) 寫出完整 spec；有草稿、截圖或「找圖」需求時先呼叫 collect_assets 確認素材齊全（缺的請使用者補）；"
    u"3) build_deck；4) preview_deck 看每頁圖片，用 patch_deck 只修有問題的頁（最多 2 輪）；"
    u"5) review_deck 做整份的配色／字型／對齊／節奏總檢；6) 回報使用者。"
    u"主管在 PowerPoint 留註解後用 read_feedback 讀取再 patch_deck（自動另存新版本）。"
    u"使用者滿意的成品可用 save_example 存成範例；要學公司舊簡報的配色字型用 import_style。數字一律用 {t1…} 引用。"
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
                "pdf": {"type": "boolean", "description": u"同時用本機 PowerPoint 匯出 PDF"},
            },
            "required": ["spec"],
        },
    },
    {
        "name": "patch_deck",
        "description": (u"只修改指定頁再重新產出（其他頁不動）。頁碼是 spec 的頁碼（build_deck 回傳的對照表左欄）。"
                        u"會另存新版本（_v2、_v3…）並附變更紀錄，不覆蓋原檔；原檔在 PowerPoint 手動改過會列出差異。"),
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
                "pdf": {"type": "boolean", "description": u"同時匯出 PDF"},
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
        "name": "collect_assets",
        "description": (
            u"動工前的確認：不產生正式簡報。回傳版面計畫表、每頁線框圖（spec 有 sketch 時和手繪草稿並排，"
            u"讓使用者確認版面與註解判讀）、待補素材清單，並在素材資料夾搜尋草稿上「找圖」的圖片（附編號縮圖）。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "spec": {"type": "object", "additionalProperties": True,
                         "description": u"完整 spec；還沒有的圖片寫 {\"need\":\"OpenProject 甘特圖截圖\",\"search\":[\"甘特\",\"gantt\"]}"},
                "folders": {"type": "array", "items": _STR, "description": u"要搜尋的素材資料夾；省略時用 spec.asset_dirs、assets/、screenshots/ 等"},
                "slides": {"type": "string", "description": u"線框圖只看這些投影片，例如 9-16（一次最多 8 張）"},
            },
            "required": ["spec"],
        },
    },
    {
        "name": "review_deck",
        "description": (
            u"完稿前的美術總檢（整份簡報）：配色是否都在品牌色票內、字型、標題與內容對齊、連續頁版面是否單調、"
            u"每頁顏色是否太花，並附上所有投影片的縮圖總覽（需本機 PowerPoint）。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string", "description": u".pptx 或 .spec.json 路徑"}},
            "required": ["path"],
        },
    },
    {
        "name": "read_feedback",
        "description": (
            u"讀取主管在 PowerPoint 留的註解（新式與舊式註解都支援），對應回 spec 頁碼，"
            u"並列出檔案在 PowerPoint 裡被手動改過的地方。之後把未解決的註解合併成一次 patch_deck。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {"path": {"type": "string", "description": u"主管改過／留了註解的 .pptx（或它的 .spec.json）"}},
            "required": ["path"],
        },
    },
    {
        "name": "import_style",
        "description": (
            u"從公司現有的好簡報（.pptx/.potx）抽出品牌主色、輔色、灰階、字型與 Logo，回傳設定與色票圖；"
            u"使用者確認後用 apply=true 寫入 design_system/theme.yaml（自動備份舊檔）。"
        ),
        "inputSchema": {
            "type": "object",
            "properties": {"pptx_path": _STR, "apply": {"type": "boolean", "description": u"true＝寫入 theme.yaml"}},
            "required": ["pptx_path"],
        },
    },
    {
        "name": "save_example",
        "description": u"把使用者滿意的成品存成範例（design_system/examples/），之後 get_design_context 會優先參考同受眾最新的範例。",
        "inputSchema": {
            "type": "object",
            "properties": {"spec_path": {"type": "string", "description": u"成品的 .spec.json"},
                           "name": {"type": "string", "description": u"範例名稱，例如「Q3月報」"}},
            "required": ["spec_path", "name"],
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
    lines = [u"✅ 已產出：%s（%d 張投影片）" % (res["path"], res["slide_count"])]
    if res.get("previous"):
        lines.append(u"（另存新版本，原檔 %s 保留不動）" % res["previous"])
    if res.get("pdf"):
        lines.append(u"PDF：%s" % res["pdf"])
    lines += [u"spec：%s（之後用 patch_deck 修改）" % res["spec_path"],
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
    if res.get("changes"):
        lines += ["", u"### 變更紀錄（第 %d 版）" % deck._version_of(res["path"])]
        lines += [u"- " + c for c in res["changes"]]
        lines.append(u"完整紀錄：%s" % res.get("changelog", ""))
    if res.get("manual_edits"):
        from .feedback import format_edits
        lines += ["", format_edits(res["manual_edits"])]
    if res.get("assets_needed"):
        lines += ["", u"### 待補素材（%d）——投影片上放了待補框" % len(res["assets_needed"])]
        lines += [u"- %s：%s" % (a["where"], a["need"]) for a in res["assets_needed"]]
        lines.append(u"可用 collect_assets 搜尋素材資料夾，或請使用者提供檔案後用 patch_deck 補上 path。")
    if res.get("lint"):
        lines += ["", u"### " + format_lint(res["lint"])]
    lines += ["", u"下一步：呼叫 preview_deck(path=\"%s\") 看每頁圖片自我檢查（對照 get_design_context 的視覺設計指引），"
                  u"只針對 ⚠ 項目與圖上看得到的問題用 patch_deck 修正，最多 2 輪；完成後呼叫 review_deck 做整份美術總檢，"
                  u"再把檔案路徑、註解解讀表與重要提醒簡短告訴使用者。" % res["path"]]
    return "\n".join(lines)


def h_get_design_context(a):
    text, pages = build_context(a.get("audience"), _obj(a.get("sources")) or [], a.get("sketch"),
                                a.get("sketch_pages"), a.get("blueprint"), a.get("template"))
    return {"content": [{"type": "text", "text": text}] + _images(pages), "isError": False}


def h_build_deck(a):
    res = deck.build(_obj(a.get("spec")), a.get("out_path"), pdf=bool(a.get("pdf")))
    return _text(format_build(res), not res.get("ok"))


def h_patch_deck(a):
    res = deck.patch(a["spec_path"], _obj(a.get("slides")), _obj(a.get("delete")), _obj(a.get("insert")),
                     a.get("out_path"), pdf=bool(a.get("pdf")))
    return _text(format_build(res), not res.get("ok"))


def h_preview_deck(a):
    p = _deck_path(a["path"])
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


def h_collect_assets(a):
    from .assets import collect

    text, imgs = collect(_obj(a.get("spec")) or {}, _obj(a.get("folders")), slides=a.get("slides"))
    return {"content": [{"type": "text", "text": text}] + imgs, "isError": False}


def _deck_path(raw):
    p = resolve_path(raw)
    if p.suffix.lower() == ".json":
        out = (C.read_json(p, {}) or {}).get("_output")
        p = resolve_path(out) if out else p.with_suffix(".pptx")
    return p


def h_review_deck(a):
    from .review import format_review, overview_sheet, review_deck

    p = _deck_path(a["path"])
    if not p.exists():
        return _text(u"找不到簡報：%s" % a["path"], True)
    res = review_deck(p)
    text = format_review(res)
    sheet, err = overview_sheet(p, res["slide_count"])
    content = []
    if sheet:
        text += u"\n\n下方是整份縮圖總覽：從遠處看整體——配色是否統一、每頁重心是否清楚、節奏是否有變化、" \
                u"截圖與圖表風格是否一致。只修 ⚠ 項目與總覽中明顯不協調的頁（用 patch_deck）。"
        content = [{"type": "image", "data": sheet, "mimeType": "image/jpeg"}]
    else:
        text += u"\n\n（沒有縮圖總覽：%s）" % (err or u"未知原因")
    return {"content": [{"type": "text", "text": text}] + content, "isError": False}


def h_read_feedback(a):
    from .feedback import format_feedback, manual_edits, read_comments, slide_texts
    from .template import open_presentation

    raw = resolve_path(a["path"])
    p = _deck_path(a["path"])
    if not p.exists():
        return _text(u"找不到簡報：%s" % a["path"], True)
    spec_file = raw if raw.suffix.lower() == ".json" else p.with_suffix(".spec.json")
    info = (C.read_json(spec_file, {}) or {}).get("_build") if spec_file.exists() else None
    comments = read_comments(p)
    titles = [t.split("\n", 1)[0] for t in slide_texts(open_presentation(p))]
    text = format_feedback(p, comments, info, manual_edits(info, p), titles)
    if spec_file.exists():
        text += u"\n\npatch_deck 的 spec_path：%s" % spec_file
    return _text(text)


def h_import_style(a):
    from .style_import import import_style

    p = resolve_path(a["pptx_path"])
    if not p.exists():
        return _text(u"找不到簡報：%s" % a["pptx_path"], True)
    text, img = import_style(p, apply=bool(a.get("apply")))
    return {"content": [{"type": "text", "text": text}, {"type": "image", "data": img, "mimeType": "image/jpeg"}],
            "isError": False}


def h_save_example(a):
    try:
        dest = C.save_example(resolve_path(a["spec_path"]), a.get("name") or "example")
    except ValueError as e:
        return _text(str(e), True)
    return _text(u"已存成範例：%s。之後同受眾做報告時，get_design_context 會優先參考最新的 2 份範例。" % dest)


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
    "import_style": h_import_style,
    "save_example": h_save_example,
    "read_feedback": h_read_feedback,
    "collect_assets": h_collect_assets,
    "review_deck": h_review_deck,
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
