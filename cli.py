#!/usr/bin/env python
"""Command-line access to the same engine (for when MCP is not allowed, or for testing).

  python cli.py context --audience 主管 --sources data/Q3.xlsx [--sketch sketches/q3.pdf]
  python cli.py build spec.json [-o output/]
  python cli.py patch output/x.spec.json patch.json
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
