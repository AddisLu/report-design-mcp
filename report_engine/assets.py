"""collect_assets: before building, check every region has its material (data, screenshots, photos) and
search the asset folders for images the sketch asked for ("找圖：…")."""
from __future__ import annotations

import base64
import io
import os
import re
from pathlib import Path

from PIL import Image, ImageDraw

from .paths import ROOT, resolve_path
from .renderer import plain_text

IMG_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff", ".svg")
MAX_FILES = 5000
TYPE_NAMES = {
    "text": u"文字／條列", "headline": u"結論句", "note": u"備註", "kpis": "KPI", "status": u"狀態燈",
    "chart": u"圖表", "table": u"表格", "gantt": u"甘特圖", "timeline": u"時程表", "process": u"流程",
    "svg": u"示意圖", "image": u"圖片", "images": u"圖組", "placeholder": u"⚠待補圖片",
}
_ASCII = re.compile(r"[a-z0-9]{2,}")
_CJK = re.compile(u"[一-鿿]+")


def tokens(text):
    s = str(text or "").lower()
    out = set(_ASCII.findall(s))
    for run in _CJK.findall(s):
        if len(run) == 1:
            out.add(run)
        out.update(run[i:i + 2] for i in range(len(run) - 1))
    return out


def default_folders(spec, base_dir=None):
    cands = []
    for d in spec.get("asset_dirs") or []:
        cands.append(resolve_path(d, base_dir))
    for name in ("assets", "images", "screenshots", u"截圖", u"圖片", u"素材"):
        cands += [Path.cwd() / name, ROOT / name]
    for src in spec.get("sources") or []:
        p = resolve_path(src, base_dir)
        cands.append(p if p.is_dir() else p.parent)
    seen, out = set(), []
    for c in cands:
        try:
            c = c.resolve()
        except OSError:
            continue
        if c.is_dir() and str(c) not in seen:
            seen.add(str(c))
            out.append(c)
    return out


def list_images(folders):
    files = []
    for folder in folders:
        for root, dirs, names in os.walk(str(folder)):
            dirs[:] = [d for d in dirs if not d.startswith((".", "_", "~")) and d not in ("output", ".venv", "wheelhouse")]
            for n in names:
                if n.lower().endswith(IMG_EXT) and not n.startswith("~$"):
                    files.append(Path(root) / n)
                    if len(files) >= MAX_FILES:
                        return files
    return files


def search(files, need, keywords, limit=6):
    query = tokens(need)
    for k in keywords or []:
        query |= tokens(k)
        query.add(str(k).lower())
    scored = []
    for f in files:
        name, folder = f.stem.lower(), str(f.parent).lower()
        hits = sum(1 for q in query if q in name)
        if hits:  # the folder name only breaks ties (a 截圖/ folder would otherwise match everything)
            scored.append((hits * 2 + sum(1 for q in query if q in folder), f.stat().st_mtime if f.exists() else 0, f))
    scored.sort(key=lambda t: (-t[0], -t[1]))
    return [f for _s, _m, f in scored[:limit]]


def contact_sheet(items, cols=3, cell=(320, 210)):
    """items: [(label, path)] -> base64 JPEG with numbered thumbnails."""
    if not items:
        return None
    rows = -(-len(items) // cols)
    cw, ch = cell
    sheet = Image.new("RGB", (cols * cw, rows * (ch + 22)), "white")
    d = ImageDraw.Draw(sheet)
    for i, (label, path) in enumerate(items):
        x, y = (i % cols) * cw, (i // cols) * (ch + 22)
        try:
            if str(path).lower().endswith(".svg"):
                raise ValueError
            with Image.open(str(path)) as im:
                im = im.convert("RGB")
                im.thumbnail((cw - 12, ch - 12))
                sheet.paste(im, (x + (cw - im.width) // 2, y + (ch - im.height) // 2))
        except Exception:
            d.rectangle([x + 6, y + 6, x + cw - 6, y + ch - 6], outline="#999")
            d.text((x + 14, y + ch // 2), "SVG / preview n/a", fill="#666")
        d.rectangle([x + 6, y + 6, x + 46, y + 26], fill="#C62828")
        d.text((x + 12, y + 9), label, fill="white")
        d.text((x + 8, y + ch + 4), Path(str(path)).name[:44], fill="#333")
    buf = io.BytesIO()
    sheet.save(buf, format="JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _blocks(ns):
    if ns["mode"] == "section":
        return []
    if ns["mode"] == "custom":
        return [r["block"] for r in ns["regions"]]
    return ns["blocks"]


def _block_label(bk):
    t = bk["type"]
    name = TYPE_NAMES.get(t, t)
    if t == "chart":
        name += "(%s)" % bk.get("chart_type")
    elif t == "images":
        name += "(%d)" % len(bk["items"])
    elif t == "image":
        try:
            with Image.open(bk["path"]) as im:
                name += u"：%s %d×%d" % (Path(bk["path"]).name, im.size[0], im.size[1])
        except Exception:
            name += u"：%s" % Path(bk["path"]).name
    return name


def _wireframe_images(spec, base_dir, slides):
    """Trial build in a temp folder, then draw what each slide will look like (layout boxes + note readings)."""
    import shutil
    import tempfile

    from .deck import build
    from .sketch import parse_pages
    from .wireframe import wireframes

    tmp = tempfile.mkdtemp(prefix="report_wire_")
    try:
        res = build(dict(spec), os.path.join(tmp, "plan.pptx"), base_dir=base_dir)
        if not res.get("ok"):
            return [], 0
        wanted = [i + 1 for i in parse_pages(slides, res["slide_count"])] if slides else None
        imgs = wireframes(res["path"], spec, res["mapping"], wanted)
        return imgs, res["slide_count"]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def collect(spec, folders=None, base_dir=None, slides=None):
    """-> (markdown text, [image dicts])."""
    from .deck import prepare

    b, nslides = prepare(spec, base_dir)
    lines = [u"# 素材清單與版面計畫", u"", u"| 頁 | 標題 | 版面區塊 | 狀態 |", u"|---|---|---|---|"]
    for ns in nslides:
        blocks = _blocks(ns)
        page = "%s%s" % (ns["idx"], u"（%s）" % ns["label"] if ns.get("label") else "")
        kind = u"章節頁" if ns["mode"] == "section" else (u"、".join(_block_label(bk) for bk in blocks) or u"封面／標題")
        missing = [bk for bk in blocks if bk["type"] == "placeholder"]
        state = u"⚠ 缺 %d 項素材" % len(missing) if missing else u"✓ 齊全"
        lines.append(u"| %s | %s | %s | %s |" % (page, plain_text(ns.get("title") or "")[:28], kind, state))
    images = []
    if b.errors:
        lines += ["", u"## ❌ 資料或寫法問題（%d）" % len(b.errors)] + [u"- " + e for e in b.errors]
    else:
        try:
            wires, total = _wireframe_images(spec, base_dir, slides)
        except Exception as e:  # the plan table is still useful without pictures
            wires, total = [], 0
            lines.append(u"\n（線框圖產生失敗：%s）" % e)
        if wires:
            lines += ["", u"## 版面判讀確認（%d 張線框圖%s）" % (len(wires), u"，共 %d 頁，可用 slides 參數看其他頁" % total
                                                         if total > len(wires) else ""),
                      u"下方先是線框圖（%s）：每個框是一個區塊與它的類型，紅字是手寫註解的解讀。"
                      % (u"左邊是手繪草稿、右邊是判讀結果" if spec.get("sketch") else u"spec 沒有寫 sketch，所以只有判讀結果"),
                      u"請使用者確認框的位置、類型與註解解讀都對，再補素材。"]
            images += [{"type": "image", "data": w, "mimeType": "image/jpeg"} for w in wires]
    if b.assets_needed:
        dirs = [resolve_path(f, base_dir) for f in folders] if folders else default_folders(spec, base_dir)
        files = list_images(dirs)
        lines += ["", u"## ⚠ 待補素材（%d）" % len(b.assets_needed),
                  u"搜尋資料夾：%s（共 %d 張圖）" % (u"、".join(str(d) for d in dirs) or u"（沒有，請用 folders 指定）", len(files))]
        sheet_items = []
        for i, a in enumerate(b.assets_needed, 1):
            cands = search(files, a["need"], a["search"])
            lines.append(u"### %d. %s —— %s" % (i, a["where"], a["need"]))
            if cands:
                for j, f in enumerate(cands, 1):
                    label = "%d-%d" % (i, j)
                    sheet_items.append((label, f))
                    lines.append(u"- [%s] %s" % (label, f))
            else:
                lines.append(u"- 沒找到候選圖：請使用者提供檔案（例如 OpenProject 畫面請用 PrintScreen 存成 PNG）")
        sheet = contact_sheet(sheet_items)
        if sheet:
            images.append({"type": "image", "data": sheet, "mimeType": "image/jpeg"})
            lines.append(u"\n下方是候選圖縮圖（左上角編號對應上面清單）。")
    if b.warnings:
        lines += ["", u"## 提醒"] + [u"- " + w for w in b.warnings[:15]]
    ready = not b.errors and not b.assets_needed
    lines += ["", u"**下一步**：先請使用者對照線框圖確認版面判讀與註解解讀（讀錯就改 spec 再呼叫一次 collect_assets）；"
              + (u"素材齊全，確認後就可以 build_deck。" if ready else
                 u"再把待補清單交給使用者：選好的候選圖填進該區塊的 path（截圖建議 style:\"browser\"），"
                 u"缺的請使用者提供；也可以先 build_deck，待補的位置會放待補框。")]
    return "\n".join(lines), images
