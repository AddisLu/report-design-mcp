#!/usr/bin/env python
"""Command-line access to the same engine (for when MCP is not allowed, or for testing).

  python cli.py context --audience 主管 --sources data/Q3.xlsx [--sketch sketches/q3.pdf]
  python cli.py build spec.json [-o output/]
  python cli.py patch output/x.spec.json patch.json
  python cli.py preview output/x.pptx [--slides 1-6] [--out output/preview]
  python cli.py assets spec.json [--folders assets screenshots] [--out output/candidates.jpg]
  python cli.py review output/x.pptx [--out output/overview.jpg]
  python cli.py feedback output/x.pptx            (read PowerPoint comments + manual edits)
  python cli.py style 公司舊簡報.pptx [--apply]     (learn palette / fonts / logo into theme.yaml)
  python cli.py example output/x.spec.json Q3月報   (save an approved deck as an example)
  python cli.py pdf output/x.pptx
  python cli.py doctor
  python cli.py canvas sketches/canvas.pdf --pages 6
  python cli.py inspect design_system/templates/company.pptx --set-default
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from report_engine import deck, tools  # noqa: E402
from report_engine.context import build_context  # noqa: E402
from report_engine.sketch import make_canvas  # noqa: E402


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="ReportDesign CLI")
    sub = ap.add_subparsers(dest="cmd")

    c = sub.add_parser("context", help="print the design context (what the model reads)")
    c.add_argument("--audience")
    c.add_argument("--sources", nargs="*", default=[])
    c.add_argument("--sketch")
    c.add_argument("--pages")
    c.add_argument("--blueprint")
    c.add_argument("--template")
    c.add_argument("--dump-images", help="folder to save the sketch page images that would be sent")

    b = sub.add_parser("build", help="spec.json -> .pptx")
    b.add_argument("spec")
    b.add_argument("-o", "--out")

    p = sub.add_parser("patch", help="apply patch.json ({slides, delete, insert}) to a saved spec")
    p.add_argument("spec_path")
    p.add_argument("patch")

    pv = sub.add_parser("preview", help="layout lint + slide images (images need PowerPoint on Windows)")
    pv.add_argument("pptx")
    pv.add_argument("--slides")
    pv.add_argument("--out", help="folder to save the slide images")

    sa = sub.add_parser("assets", help="layout plan + missing material + candidate images (before building)")
    sa.add_argument("spec")
    sa.add_argument("--folders", nargs="*")
    sa.add_argument("--out", help="save the candidate contact sheet here")

    rv = sub.add_parser("review", help="whole-deck art review (palette, fonts, alignment, rhythm)")
    rv.add_argument("pptx")
    rv.add_argument("--out", help="save the thumbnail overview here (needs PowerPoint on Windows)")

    fb = sub.add_parser("feedback", help="read PowerPoint comments and manual edits")
    fb.add_argument("pptx")
    st = sub.add_parser("style", help="learn brand palette/fonts/logo from a company deck")
    st.add_argument("pptx")
    st.add_argument("--apply", action="store_true")
    st.add_argument("--out", help="save the swatch image here")
    ex = sub.add_parser("example", help="save an approved deck's spec as an example")
    ex.add_argument("spec_path")
    ex.add_argument("name")
    pd = sub.add_parser("pdf", help="export a deck to PDF with PowerPoint")
    pd.add_argument("pptx")
    sub.add_parser("doctor", help="check Python, packages, fonts, PowerPoint and the design system")

    v = sub.add_parser("canvas", help="make a 12x8 grid canvas for the drawing tablet")
    v.add_argument("out")
    v.add_argument("--pages", type=int, default=6)
    v.add_argument("--no-requirements", action="store_true")

    i = sub.add_parser("inspect", help="read a template's layouts")
    i.add_argument("template")
    i.add_argument("--set-default", action="store_true")

    a = ap.parse_args(argv)
    if a.cmd == "context":
        text, pages = build_context(a.audience, a.sources, a.sketch, a.pages, a.blueprint, a.template)
        print(text)
        if a.dump_images:
            d = Path(a.dump_images)
            d.mkdir(parents=True, exist_ok=True)
            for pg in pages:
                (d / ("page_%02d.jpg" % pg["page"])).write_bytes(base64.b64decode(pg["data"]))
            print("\n[%d page images saved to %s]" % (len(pages), d))
        return 0
    if a.cmd == "build":
        spec_file = Path(a.spec)
        spec = json.loads(spec_file.read_text(encoding="utf-8"))
        res = deck.build(spec, a.out, base_dir=spec_file.resolve().parent)
        print(tools.format_build(res))
        return 0 if res.get("ok") else 1
    if a.cmd == "patch":
        pt = json.loads(Path(a.patch).read_text(encoding="utf-8"))
        res = deck.patch(a.spec_path, pt.get("slides"), pt.get("delete"), pt.get("insert"), pt.get("out_path"))
        print(tools.format_build(res))
        return 0 if res.get("ok") else 1
    if a.cmd == "preview":
        from report_engine import preview
        from report_engine.lint import format_lint, lint_deck

        res = lint_deck(a.pptx)
        print(format_lint(res))
        imgs, err = preview.export_images(a.pptx, a.slides, res["slide_count"])
        if imgs and a.out:
            d = Path(a.out)
            d.mkdir(parents=True, exist_ok=True)
            for im in imgs:
                (d / ("slide_%02d.jpg" % im["slide"])).write_bytes(base64.b64decode(im["data"]))
            print("\n[%d slide images saved to %s]" % (len(imgs), d))
        elif not imgs:
            print("\n[no images: %s]" % err)
        return 0
    if a.cmd == "assets":
        from report_engine.assets import collect

        spec_file = Path(a.spec)
        text, imgs = collect(json.loads(spec_file.read_text(encoding="utf-8")), a.folders,
                             base_dir=spec_file.resolve().parent)
        print(text)
        if imgs and a.out:
            Path(a.out).parent.mkdir(parents=True, exist_ok=True)
            Path(a.out).write_bytes(base64.b64decode(imgs[0]["data"]))
            print("\n[candidate sheet saved to %s]" % a.out)
        return 0
    if a.cmd == "review":
        from report_engine.review import format_review, overview_sheet, review_deck

        res = review_deck(a.pptx)
        print(format_review(res))
        sheet, err = overview_sheet(a.pptx, res["slide_count"])
        if sheet and a.out:
            Path(a.out).parent.mkdir(parents=True, exist_ok=True)
            Path(a.out).write_bytes(base64.b64decode(sheet))
            print("\n[overview saved to %s]" % a.out)
        elif not sheet:
            print("\n[no overview: %s]" % err)
        return 0
    if a.cmd == "feedback":
        r = tools.h_read_feedback({"path": a.pptx})
        print(r["content"][0]["text"])
        return 1 if r["isError"] else 0
    if a.cmd == "style":
        from report_engine.style_import import import_style

        text, img = import_style(a.pptx, apply=a.apply)
        print(text)
        if a.out:
            Path(a.out).write_bytes(base64.b64decode(img))
        return 0
    if a.cmd == "example":
        r = tools.h_save_example({"spec_path": a.spec_path, "name": a.name})
        print(r["content"][0]["text"])
        return 1 if r["isError"] else 0
    if a.cmd == "pdf":
        from report_engine.preview import export_pdf

        path, err = export_pdf(a.pptx)
        print(path or err)
        return 0 if path else 1
    if a.cmd == "doctor":
        from report_engine.doctor import report

        print(report())
        return 0
    if a.cmd == "canvas":
        for path in make_canvas(a.out, a.pages, not a.no_requirements):
            print(path)
        return 0
    if a.cmd == "inspect":
        r = tools.h_inspect_template({"template_path": a.template, "set_default": a.set_default})
        print(r["content"][0]["text"])
        return 1 if r["isError"] else 0
    ap.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
