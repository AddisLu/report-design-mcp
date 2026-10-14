"""Generate a native 16:9 brand template from design tokens (used when no company template is provided).

Layouts: 封面 / 章節 / 標題及內容 / 兩欄 / 僅標題 / 空白. Theme colours and fonts, master text styles,
footer + slide number, and an optional logo all come from design_system/theme.yaml.
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu

from .theme import contrast, mix

BRAND_VERSION = "1"
W, H = 12192000, 6858000
IN = 914400
_NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
KEEP = (  # (default layout name, new name)
    ("Title Slide", u"封面"),
    ("Section Header", u"章節"),
    ("Title and Content", u"標題及內容"),
    ("Two Content", u"兩欄"),
    ("Title Only", u"僅標題"),
    ("Blank", u"空白"),
)


def _in(v):
    return int(round(v * IN))


class _Geo:
    def __init__(self, theme):
        self.mx = theme.inch("margin_x") or _in(0.6)
        self.mt = theme.inch("margin_top") or _in(0.4)
        self.cw = W - 2 * self.mx
        self.title_h = _in(0.85)
        self.body_top = self.mt + self.title_h + _in(0.2)
        self.foot_y = H - _in(0.48)
        self.foot_h = _in(0.3)
        self.body_h = self.foot_y - _in(0.15) - self.body_top
        self.gap = theme.inch("gap") or _in(0.2)


def _xfrm(x, y, w, h):
    return '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (x, y, w, h)


def _clr(hex6):
    return '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % hex6


def _ph(sid, name, ph, rect, body_pr="", lst="", text=""):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr %s/><a:lstStyle>%s</a:lstStyle><a:p>%s<a:endParaRPr lang="zh-TW"/></a:p></p:txBody></p:sp>'
        % (sid, name, ph, _xfrm(*rect), body_pr, lst, text)
    )


def _rect(sid, name, rect, fill_hex):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>%s<a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
        % (sid, name, _xfrm(*rect), _clr(fill_hex))
    )


_SLDNUM_FLD = ('<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
               '<a:rPr lang="zh-TW"/><a:t>&#8249;#&#8250;</a:t></a:fld>')


def _footer_phs(g, theme, start_id, master=False):
    cap = int(theme.size("caption") * 100)
    lst_l = '<a:lvl1pPr algn="l"><a:defRPr sz="%d">%s</a:defRPr></a:lvl1pPr>' % (cap, _clr(theme.neutral(500)))
    lst_r = '<a:lvl1pPr algn="r"><a:defRPr sz="%d">%s</a:defRPr></a:lvl1pPr>' % (cap, _clr(theme.neutral(500)))
    ftr_idx, num_idx = (3, 4) if master else (11, 12)
    body = 'lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"'
    return (
        _ph(start_id, "Footer", '<p:ph type="ftr" sz="quarter" idx="%d"/>' % ftr_idx,
            (g.mx, g.foot_y, int(g.cw * 0.7), g.foot_h), body, lst_l)
        + _ph(start_id + 1, "Slide Number", '<p:ph type="sldNum" sz="quarter" idx="%d"/>' % num_idx,
              (W - g.mx - _in(1.0), g.foot_y, _in(1.0), g.foot_h), body, lst_r, _SLDNUM_FLD)
    )


def _run_style(size, hex6=None, bold=False, font="+mn"):
    return ('<a:defRPr sz="%d"%s kern="1200">%s<a:latin typeface="%s-lt"/><a:ea typeface="%s-ea"/>'
            '<a:cs typeface="%s-cs"/></a:defRPr>'
            % (int(size * 100), ' b="1"' if bold else ' b="0"',
               _clr(hex6) if hex6 else '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>', font, font, font))


def _tx_styles(theme):
    t = theme
    title = ('<p:titleStyle><a:lvl1pPr algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
             '<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
             '%s</a:lvl1pPr></p:titleStyle>' % _run_style(t.size("title"), bold=True, font="+mj"))
    lv = []
    ind = _in(0.25)
    for i in range(9):
        k = min(i, 2)
        color = None if k == 0 else t.neutral(700)
        bu = u"•" if k != 1 else u"–"
        lv.append(
            '<a:lvl%dpPr marL="%d" indent="%d" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
            '<a:lnSpc><a:spcPct val="115000"/></a:lnSpc><a:spcBef><a:spcPts val="%d"/></a:spcBef>'
            '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="%s"/>'
            '%s</a:lvl%dpPr>' % (i + 1, ind * (i + 1), -ind, 800 if k == 0 else 400, bu,
                                 _run_style(t.size("body", k), color), i + 1))
    body = "<p:bodyStyle>%s</p:bodyStyle>" % "".join(lv)
    other = ('<p:otherStyle><a:defPPr><a:defRPr lang="zh-TW"/></a:defPPr>'
             '<a:lvl1pPr marL="0" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">%s</a:lvl1pPr>'
             '</p:otherStyle>' % _run_style(t.size("body", 1)))
    return parse_xml('<p:txStyles %s>%s%s%s</p:txStyles>' % (_NS, title, body, other))


def _replace_tree(el, shapes_xml, bg_hex=None, name=None):
    """Replace a master/layout <p:cSld> with our own shapes (keeps the mandatory group header)."""
    csld = el.find(qn("p:cSld"))
    if name is not None:
        csld.set("name", name)
    old_bg = csld.find(qn("p:bg"))
    if old_bg is not None:
        csld.remove(old_bg)
    if bg_hex:
        csld.insert(0, parse_xml('<p:bg %s><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>' % (_NS, _clr(bg_hex))))
    elif el.tag == qn("p:sldMaster"):
        csld.insert(0, parse_xml('<p:bg %s><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' % _NS))
    tree = csld.find(qn("p:spTree"))
    for child in list(tree)[2:]:
        tree.remove(child)
    frag = parse_xml('<p:spTree %s>%s</p:spTree>' % (_NS, shapes_xml))
    for child in list(frag):
        tree.append(child)
    return tree


def _theme_xml(blob, theme):
    root = etree.fromstring(blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    cs = root.find(".//%sclrScheme" % a)
    cs.set("name", "ReportDesign")
    accents = [theme.primary, theme.secondary] + [c for c in theme.chart_colors if c not in (theme.primary, theme.secondary)]
    accents = (accents + theme.chart_colors * 3)[:6]
    colors = {"dk1": theme.neutral(900), "lt1": "FFFFFF", "dk2": theme.neutral(700), "lt2": theme.neutral(50),
              "hlink": theme.primary, "folHlink": theme.neutral(500)}
    for i, c in enumerate(accents, 1):
        colors["accent%d" % i] = c
    for tag, val in colors.items():
        el = cs.find(a + tag)
        if el is None:
            continue
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, a + "srgbClr").set("val", val)
    fs = root.find(".//%sfontScheme" % a)
    fs.set("name", "ReportDesign")
    for which in ("majorFont", "minorFont"):
        f = fs.find(a + which)
        f.find(a + "latin").set("typeface", theme.font_latin)
        ea = f.find(a + "ea")
        if ea is not None:
            ea.set("typeface", theme.font_ea)
        for sub in f.findall(a + "font"):  # per-script overrides would beat our ea font for zh-TW
            if sub.get("script") in ("Hant", "Hans", "Jpan", "Hang"):
                sub.set("typeface", theme.font_ea)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _set_theme(master, theme):
    part = master.part.part_related_by(RT.THEME)
    blob = part._element if hasattr(part, "_element") else None
    new = _theme_xml(etree.tostring(blob) if blob is not None else part.blob, theme)
    if hasattr(part, "_element"):
        part._element = parse_xml(new)
    else:
        part._blob = new


def _logo_xml(part, theme, sid, x_right, y):
    path = theme.logo_path()
    if not path:
        return "", 0
    from PIL import Image

    with Image.open(str(path)) as im:
        iw, ih = im.size
    h = theme.logo_height
    w = int(h * iw / float(ih))
    _img, rid = part.get_or_add_image_part(str(path))
    xml = ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="Logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
           '<p:nvPr userDrawn="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
           '</p:blipFill><p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
           % (sid, rid, _xfrm(x_right - w, y, w, h)))
    return xml, w


def _title_ph(g, sid=2, width=None, ph='<p:ph type="title"/>'):
    return _ph(sid, "Title", ph, (g.mx, g.mt, width or g.cw, g.title_h), 'lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"')


def _body_ph(sid, idx, rect):
    return _ph(sid, "Content %d" % idx, '<p:ph idx="%d"/>' % idx, rect, 'lIns="0" rIns="0" tIns="0" bIns="0"')


def build_brand_template(dest: Path, theme) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(W), Emu(H)
    master = prs.slide_masters[0]
    g = _Geo(theme)
    _set_theme(master, theme)

    # master: accent tick above the title, title, body, footer, slide number, logo
    logo, logo_w = _logo_xml(master.part, theme, 20, W - g.mx, g.mt - _in(0.02))
    title_w = g.cw - (logo_w + _in(0.3) if logo_w else 0)
    shapes = (_rect(10, "Accent", (g.mx, g.mt - _in(0.14), _in(0.45), _in(0.05)), theme.primary)
              + _title_ph(g, 2, title_w)
              + _ph(3, "Body", '<p:ph type="body" idx="1"/>', (g.mx, g.body_top, g.cw, g.body_h),
                    'lIns="0" rIns="0" tIns="0" bIns="0"')
              + _footer_phs(g, theme, 4, master=True) + logo)
    _replace_tree(master._element, shapes)
    old = master._element.find(qn("p:txStyles"))
    master._element.replace(old, _tx_styles(theme))

    keep = dict(KEEP)
    for layout in list(master.slide_layouts):
        if layout.name not in keep:
            master.slide_layouts.remove(layout)

    on_primary = "FFFFFF" if contrast("FFFFFF", theme.primary) >= 3.0 else theme.neutral(900)
    sub_on_primary = mix(theme.primary, on_primary, 0.8)
    half = (g.cw - g.gap) // 2
    for layout in master.slide_layouts:
        name = layout.name
        el = layout._element
        bg = None
        if name == "Title Slide":
            bg = theme.primary
            cover_logo, _w = _logo_xml(layout.part, theme, 20, W - g.mx, H - g.mt - theme.logo_height)
            shapes = (
                _rect(10, "Accent", (_in(0.9), _in(3.95), _in(0.8), _in(0.06)), theme.secondary)
                + _ph(2, "Title", '<p:ph type="ctrTitle"/>', (_in(0.9), _in(1.9), _in(10.5), _in(1.9)),
                      'lIns="0" rIns="0" tIns="0" bIns="0" anchor="b"',
                      '<a:lvl1pPr algn="l"><a:defRPr sz="%d" b="1">%s</a:defRPr></a:lvl1pPr>'
                      % (int(theme.size("cover") * 100), _clr(on_primary)))
                + _ph(3, "Subtitle", '<p:ph type="subTitle" idx="1"/>', (_in(0.9), _in(4.2), _in(10.5), _in(1.6)),
                      'lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"',
                      '<a:lvl1pPr marL="0" indent="0" algn="l"><a:buNone/><a:defRPr sz="%d">%s</a:defRPr></a:lvl1pPr>'
                      '<a:lvl2pPr marL="0" indent="0" algn="l"><a:buNone/><a:defRPr sz="%d">%s</a:defRPr></a:lvl2pPr>'
                      % (int(theme.size("subtitle") * 100), _clr(sub_on_primary),
                         int(theme.size("caption") * 100 + 200), _clr(sub_on_primary)))
                + cover_logo)
            el.set("showMasterSp", "0")
        elif name == "Section Header":
            bg = theme.neutral(50)
            shapes = (
                _rect(10, "Accent", (g.mx, _in(2.55), _in(0.08), _in(2.0)), theme.primary)
                + _ph(2, "Title", '<p:ph type="title"/>', (g.mx + _in(0.35), _in(2.45), _in(10.5), _in(1.3)),
                      'lIns="0" rIns="0" tIns="0" bIns="0" anchor="b"',
                      '<a:lvl1pPr algn="l"><a:defRPr sz="%d" b="1"/></a:lvl1pPr>' % int(theme.size("section") * 100))
                + _ph(3, "Subtitle", '<p:ph type="subTitle" idx="1"/>', (g.mx + _in(0.35), _in(3.9), _in(10.5), _in(0.9)),
                      'lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"',
                      '<a:lvl1pPr marL="0" indent="0" algn="l"><a:buNone/><a:defRPr sz="%d">%s</a:defRPr></a:lvl1pPr>'
                      % (int(theme.size("subtitle") * 100), _clr(theme.neutral(500)))))
            el.set("showMasterSp", "0")
        elif name == "Title and Content":
            shapes = _title_ph(g, 2, title_w) + _body_ph(3, 1, (g.mx, g.body_top, g.cw, g.body_h)) + _footer_phs(g, theme, 4)
        elif name == "Two Content":
            shapes = (_title_ph(g, 2, title_w)
                      + _body_ph(3, 1, (g.mx, g.body_top, half, g.body_h))
                      + _body_ph(4, 2, (g.mx + half + g.gap, g.body_top, half, g.body_h))
                      + _footer_phs(g, theme, 5))
        elif name == "Title Only":
            shapes = _title_ph(g, 2, title_w) + _footer_phs(g, theme, 3)
        else:  # Blank
            shapes = _footer_phs(g, theme, 2)
        _replace_tree(el, shapes, bg, keep[name])

    dest.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dest))
    return dest
