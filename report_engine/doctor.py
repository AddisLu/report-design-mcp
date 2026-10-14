"""python cli.py doctor — check the environment once after installing (or when something looks wrong)."""
from __future__ import annotations

import os
import platform
import struct
import sys
import tempfile
from pathlib import Path

FONT_FILES = {
    u"微軟正黑體（Microsoft JhengHei）": ("msjh.ttc", "msjh.ttf"),
    "Segoe UI": ("segoeui.ttf",),
}
_PS_VERSION = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
Write-Output ("VERSION=" + $app.Version)
if ($app.Presentations.Count -eq 0) { $app.Quit() }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
"""


def _font_dirs():
    dirs = [Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    return dirs


def checks():
    """-> [(ok: True/False/None, item, detail, fix)]; None = not applicable / warning."""
    out = []
    v = sys.version_info
    bits = struct.calcsize("P") * 8
    out.append((v[:2] >= (3, 8) and bits == 64, u"Python", u"%d.%d.%d，%d 位元" % (v[0], v[1], v[2], bits),
                u"請用 64 位元 Python 3.8 以上（install_offline.bat 會檢查）"))
    for mod, label in (("pptx", "python-pptx"), ("lxml", "lxml"), ("PIL", "Pillow"), ("yaml", "PyYAML"),
                       ("openpyxl", "openpyxl"), ("docx", "python-docx"), ("pypdfium2", "pypdfium2")):
        try:
            m = __import__(mod)
            out.append((True, label, getattr(m, "__version__", "OK"), ""))
        except Exception as e:  # noqa: BLE001
            out.append((False, label, u"無法載入：%s" % e, u"重新執行 install_offline.bat"))
    if os.name == "nt":
        for name, files in FONT_FILES.items():
            found = any((d / f).exists() for d in _font_dirs() for f in files)
            out.append((found, u"字型 " + name, u"已安裝" if found else u"找不到",
                        u"請安裝字型，或在 design_system/theme.yaml 的 fonts 改成已安裝的字型"))
        from .preview import run_script
        ok, stdout, err = run_script(_PS_VERSION, [], timeout=60)
        ver = next((l.split("=", 1)[1].strip() for l in stdout.splitlines() if l.startswith("VERSION=")), "")
        out.append((ok, u"PowerPoint（預覽圖、縮圖總覽、PDF）", u"版本 %s" % ver if ok else err,
                    u"沒有桌面版 PowerPoint 時仍可產出簡報，只是沒有預覽圖與 PDF"))
    else:
        out.append((None, u"字型／PowerPoint", u"%s 不是 Windows，略過" % platform.system(),
                    u"預覽圖與 PDF 需要 Windows＋桌面版 PowerPoint"))
    from .config import load_config
    from .paths import output_dir

    try:
        from .template import resolve_template_path
        from .theme import load_theme

        th = load_theme()
        tpl = resolve_template_path(None, load_config())
        out.append((True, u"設計系統", u"主色 #%s；範本 %s" % (th.primary, tpl.name), ""))
    except Exception as e:  # noqa: BLE001
        out.append((False, u"設計系統", str(e), u"檢查 design_system/theme.yaml 與 templates/"))
    try:
        cfg = load_config()
        od = Path(cfg["output_dir"]) if cfg.get("output_dir") else output_dir()
        od.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=str(od), delete=True):
            pass
        out.append((True, u"輸出資料夾", str(od), ""))
    except Exception as e:  # noqa: BLE001
        out.append((False, u"輸出資料夾", str(e), u"在 design_system/config.json 的 output_dir 指定可寫入的資料夾"))
    return out


def report() -> str:
    rows = checks()
    lines = [u"ReportDesign 環境檢查"]
    for ok, item, detail, fix in rows:
        mark = u"✓" if ok else (u"－" if ok is None else u"✗")
        lines.append(u"%s %s：%s" % (mark, item, detail))
        if ok is False and fix:
            lines.append(u"    → " + fix)
    bad = sum(1 for r in rows if r[0] is False)
    lines.append(u"\n%s" % (u"全部正常。" if not bad else u"有 %d 項需要處理。" % bad))
    return "\n".join(lines)
