"""Render slides to images with the locally installed PowerPoint (COM via PowerShell; no extra packages).

Returns None when PowerPoint is not available (non-Windows, not installed, timeout); callers then fall back
to the geometric lint report only.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

from PIL import Image

from .sketch import _encode, parse_pages

MAX_SLIDES = 12
TIMEOUT = 120

_PS = r"""
param([string]$In, [string]$Out, [string]$Slides)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
  # Open(FileName, ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse)
  $pres = $app.Presentations.Open($In, -1, 0, 0)
  $n = $pres.Slides.Count
  foreach ($s in $Slides.Split(',')) {
    $i = [int]$s
    if ($i -ge 1 -and $i -le $n) {
      $pres.Slides.Item($i).Export((Join-Path $Out ('slide_{0:D3}.png' -f $i)), 'PNG', 1600, 900)
    }
  }
  Write-Output "SLIDES=$n"
} finally {
  if ($pres -ne $null) { $pres.Close() }
  if ($app.Presentations.Count -eq 0) { $app.Quit() }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
"""


def _powershell() -> Optional[str]:
    if os.name != "nt":
        return None
    return shutil.which("powershell") or shutil.which("pwsh")


def available() -> bool:
    return _powershell() is not None


def export_images(pptx_path, slides: Optional[str] = None, total: int = 0, max_edge: int = 1280, limit: int = MAX_SLIDES):
    """-> (list of {"slide": n, "data": base64 jpeg}, error message or "")."""
    ps = _powershell()
    if ps is None:
        return None, u"這台電腦不是 Windows，無法用 PowerPoint 產生預覽圖"
    idx = [i + 1 for i in parse_pages(slides, total)] if total else []
    if not idx:
        return [], u"沒有要預覽的投影片"
    idx = idx[:limit]
    tmp = Path(tempfile.mkdtemp(prefix="report_preview_"))
    script = tmp / "export.ps1"
    script.write_text(_PS, encoding="utf-8-sig")  # BOM so Windows PowerShell 5 reads it as UTF-8
    try:
        proc = subprocess.run(
            [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script),
             "-In", str(Path(pptx_path).resolve()), "-Out", str(tmp), "-Slides", ",".join(str(i) for i in idx)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=TIMEOUT)
        if proc.returncode != 0:
            err = proc.stderr.decode("utf-8", "replace").strip().splitlines()
            return None, u"PowerPoint 匯出失敗：%s" % (err[-1] if err else proc.returncode)
        out: List[dict] = []
        for i in idx:
            png = tmp / ("slide_%03d.png" % i)
            if png.exists():
                with Image.open(str(png)) as im:
                    out.append({"slide": i, "data": _encode(im, max_edge)})
        return out, ""
    except subprocess.TimeoutExpired:
        return None, u"PowerPoint 匯出逾時（%d 秒）" % TIMEOUT
    except OSError as e:
        return None, u"無法啟動 PowerShell：%s" % e
    finally:
        shutil.rmtree(str(tmp), ignore_errors=True)
