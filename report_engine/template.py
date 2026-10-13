"""Template loading (.pptx / .potx), slide clearing, and the built-in 16:9 fallback template."""
from __future__ import annotations

import io
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu

from .config import load_config
from .paths import design_dir, resolve_path

_TEMPLATE_CT = b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
_PRES_CT = b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
TEMPLATE_EXT = (".pptx", ".potx")
BUILTIN_NAME = "_builtin_16x9.pptx"


def _potx_as_pptx(path: Path) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(str(path)) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(_TEMPLATE_CT, _PRES_CT)
            dst.writestr(item, data)
    buf.seek(0)
    return buf


def open_presentation(path):
    path = Path(path)
    if path.suffix.lower() == ".potx":
        return Presentation(_potx_as_pptx(path))
    return Presentation(str(path))


def remove_all_slides(prs) -> None:
    """Company templates often ship with sample slides; start from an empty deck."""
    sld_lst = prs.slides._sldIdLst
    for sld in list(sld_lst):
        r_id = sld.rId
        sld_lst.remove(sld)
        prs.part.drop_rel(r_id)


def _scale_x(shape, sx: float) -> None:
    sp_pr = getattr(shape._element, "spPr", None)
    if sp_pr is None or sp_pr.find(qn("a:xfrm")) is None:
        return  # position inherited from the master, which is scaled already
    shape.left = Emu(int(shape.left * sx))
    shape.width = Emu(int(shape.width * sx))


def build_builtin_template(dest: Path, font_ea: str = "") -> Path:
    """python-pptx's default template is 4:3; stretch it to 16:9 and set an East Asian theme font."""
    prs = Presentation()
    old_w = prs.slide_width
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    sx = prs.slide_width / old_w
    for master in prs.slide_masters:
        for shp in master.shapes:
            _scale_x(shp, sx)
        for layout in master.slide_layouts:
            for shp in layout.shapes:
                _scale_x(shp, sx)
        if font_ea:
            theme = master.part.part_related_by(RT.THEME)
            new = theme.blob.replace(b'<a:ea typeface=""/>', ('<a:ea typeface="%s"/>' % font_ea).encode("utf-8"))
            if hasattr(theme, "_element"):
                theme._element = parse_xml(new)
            else:
                theme._blob = new
    dest.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dest))
    return dest


def resolve_template_path(explicit=None, cfg=None) -> Path:
    cfg = cfg or load_config()
    tdir = design_dir() / "templates"
    if explicit:
        p = resolve_path(explicit, base=design_dir())
        if not p.exists():
            raise FileNotFoundError("找不到範本：%s" % explicit)
        return p
    if cfg.get("template"):
        p = resolve_path(cfg["template"], base=design_dir())
        if p.exists():
            return p
    if tdir.exists():
        cands = sorted(
            x for x in tdir.iterdir()
            if x.suffix.lower() in TEMPLATE_EXT and not x.name.startswith(("_builtin", "~$"))
        )
        if cands:
            return cands[0]
    builtin = tdir / BUILTIN_NAME
    if not builtin.exists():
        build_builtin_template(builtin, cfg.get("builtin_font_ea") or "")
    return builtin
