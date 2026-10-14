"""Hand-drawn sketches: printable/drawable 12×8 grid canvas, and loading multi-page PDF / image sketches."""
from __future__ import annotations

import base64
import io
import os
import re
from pathlib import Path
from typing import List, Optional

from PIL import Image, ImageDraw, ImageFont

from .paths import resolve_path

CANVAS_W, CANVAS_H = 1600, 1000
GRID = (80, 80, 1440, 810)  # x, y, w, h — 16:9 like the slide
IMG_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
FONT_CANDIDATES = [
    "C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/msjhl.ttc", "C:/Windows/Fonts/mingliu.ttc",
    "C:/Windows/Fonts/msyh.ttc", "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
]
GRAY = (140, 146, 156)


def _font(size):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size), True
            except OSError:
                continue
    return ImageFont.load_default(), False


def _layout_page() -> Image.Image:
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "white")
    d = ImageDraw.Draw(img)
    gx, gy, gw, gh = GRID
    small, cjk = _font(18)
    mid, _ = _font(24)
    for i in range(13):
        x = gx + gw * i / 12.0
        d.line([(x, gy), (x, gy + gh)], fill=(205, 210, 218) if i % 3 else (165, 172, 182), width=1 if i % 3 else 2)
        d.text((x - 5, gy - 26), str(i), fill=GRAY, font=small)
    for j in range(9):
        y = gy + gh * j / 8.0
        d.line([(gx, y), (gx + gw, y)], fill=(205, 210, 218) if j % 2 else (165, 172, 182), width=1 if j % 2 else 2)
        d.text((gx - 30, y - 11), str(j), fill=GRAY, font=small)
    d.rectangle([gx, gy, gx + gw, gy + gh], outline=(110, 118, 128), width=2)
    if cjk:
        head = u"類型：版面 ／ 需求 ／ ×每個______"
        page = u"頁碼 P____"
        legend = (u"黑筆＝框線＋類型：標題／結論／KPI／圖／表／重點／燈號／甘特圖／時程／流程／示意圖／截圖／照片／簡介　　"
                  u"紅筆＝註解（找圖：關鍵字、截 OpenProject…，不會印到投影片）")
    else:
        head, page, legend = "TYPE: layout / requirements / x each ____", "PAGE P____", \
            "Black pen = layout boxes (write the type inside)   Red pen = notes (never printed)"
    d.text((gx, 14), head, fill=GRAY, font=mid)
    d.text((gx + gw - 190, 14), page, fill=GRAY, font=mid)
    d.text((gx, gy + gh + 34), legend, fill=GRAY, font=small)
    return img


def _requirements_page() -> Image.Image:
    img = Image.new("RGB", (CANVAS_W, CANVAS_H), "white")
    d = ImageDraw.Draw(img)
    big, cjk = _font(34)
    small, _ = _font(22)
    if cjk:
        title = u"需求（這頁不會變成投影片）"
        fields = [u"受眾（主管／經理／工程師）", u"報告目的", u"資料來源（檔案、OpenProject 專案）",
                  u"一定要講的重點", u"素材（要貼的截圖、照片、去哪裡找）", u"其他要求（配色、風格）"]
    else:
        title = "REQUIREMENTS (not a slide)"
        fields = ["Audience", "Purpose", "Data sources", "Must-say points", "Other"]
    d.text((80, 40), title, fill=(60, 60, 60), font=big)
    y = 120
    block = (CANVAS_H - y - 40) // len(fields)
    for f in fields:
        d.text((80, y), f, fill=GRAY, font=small)
        for k in range(1, 4):
            ly = y + 30 + k * (block - 40) // 3
            d.line([(80, ly), (CANVAS_W - 80, ly)], fill=(215, 220, 226), width=1)
        y += block
    return img


def make_canvas(out_path, pages: int = 1, requirements_page: bool = False) -> List[str]:
    out = resolve_path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    imgs = ([_requirements_page()] if requirements_page else []) + [_layout_page() for _ in range(max(1, int(pages)))]
    if out.suffix.lower() == ".pdf":
        imgs[0].save(str(out), save_all=True, append_images=imgs[1:], resolution=150)
        return [str(out)]
    if out.suffix.lower() not in IMG_EXT:
        out = out.with_suffix(".png")
    if len(imgs) == 1:
        imgs[0].save(str(out))
        return [str(out)]
    paths = []
    for i, im in enumerate(imgs, 1):
        p = out.with_name("%s_%02d%s" % (out.stem, i, out.suffix))
        im.save(str(p))
        paths.append(str(p))
    return paths


def parse_pages(spec: Optional[str], total: int) -> List[int]:
    """'1-8,10' -> zero-based indices; None -> all."""
    if not spec:
        return list(range(total))
    out = []
    for part in str(spec).replace(u"，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            lo, hi = int(a or 1), int(b or total)
            out.extend(range(lo - 1, min(hi, total)))
        else:
            out.append(int(part) - 1)
    return [i for i in dict.fromkeys(out) if 0 <= i < total]


def _natural(p: Path):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", p.name)]


def _flatten(im: Image.Image) -> Image.Image:
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])  # tablet apps often export transparent backgrounds
        return bg
    return im.convert("RGB")


def _encode(im: Image.Image, max_edge: int) -> str:
    im = _flatten(im)
    im.thumbnail((max_edge, max_edge))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=82)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def load_sketch_pages(path, pages: Optional[str] = None, max_edge: int = 1280) -> List[dict]:
    """Return [{'page': n, 'label': str, 'data': base64-jpeg}] for a PDF, an image, or a folder of images."""
    p = resolve_path(path)
    if not p.exists():
        raise FileNotFoundError("找不到草稿：%s" % path)
    out = []
    if p.suffix.lower() == ".pdf":
        import pypdfium2 as pdfium

        pdf = pdfium.PdfDocument(str(p))
        try:
            for i in parse_pages(pages, len(pdf)):
                page = pdf[i]
                w, h = page.get_size()
                img = page.render(scale=max_edge / float(max(w, h))).to_pil()
                page.close()
                out.append({"page": i + 1, "label": "%s p%d" % (p.name, i + 1), "data": _encode(img, max_edge)})
        finally:
            pdf.close()
        return out
    files = sorted((x for x in p.iterdir() if x.suffix.lower() in IMG_EXT), key=_natural) if p.is_dir() else [p]
    for i in parse_pages(pages, len(files)):
        with Image.open(str(files[i])) as im:
            out.append({"page": i + 1, "label": files[i].name, "data": _encode(im, max_edge)})
    return out
