"""End-to-end smoke test. Run from the repo root:  python tests/smoke_test.py"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

TMP = tempfile.mkdtemp(prefix="report_design_test_")
DESIGN = os.path.join(TMP, "design_system")
shutil.copytree(os.path.join(ROOT, "design_system"), DESIGN)
os.environ["REPORT_DESIGN_DIR"] = DESIGN  # keep the real design_system untouched
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from pptx import Presentation  # noqa: E402

from report_engine import deck, tools  # noqa: E402
from report_engine.facts import Resolver  # noqa: E402
from report_engine.sketch import load_sketch_pages, make_canvas  # noqa: E402
from report_engine.sources import load_sources  # noqa: E402

CSV = os.path.join(ROOT, "design_system", "examples", "sample_projects.csv")
EXAMPLE = os.path.join(ROOT, "design_system", "examples", "executive_example.json")
MANAGER = os.path.join(ROOT, "design_system", "examples", "manager_example.json")
BENCH = os.path.join(ROOT, "benchmark", "mcp_reference_spec.json")


def check(cond, msg):
    if not cond:
        raise AssertionError(msg)
    print("ok -", msg)


def walk(shapes):
    for shp in shapes:
        yield shp
        if shp.shape_type == 6:  # group
            for sub in walk(shp.shapes):
                yield sub


def deck_texts(path):
    out = []
    for slide in Presentation(path).slides:
        parts = []
        for shp in walk(slide.shapes):
            if shp.has_text_frame:
                parts.append(shp.text_frame.text)
            if getattr(shp, "has_table", False) and shp.has_table:
                parts.extend(c.text for r in shp.table.rows for c in r.cells)
        out.append("\n".join(parts))
    return out


def test_facts():
    b = load_sources([CSV])
    t = b.tables["t1"]
    check(t.kinds["目標進度"] == "percent" and t.kinds["預算"] == "number", "CSV column kinds detected")
    r = Resolver(b.tables, {"gap": "=t1[代號=B].目標進度 - t1[代號=B].實際進度"})
    errs = []
    check(r.render("{t1[代號=B].實際進度}", errs) == "61%", "single cell percent")
    check(r.render("{=gap*100 | .0f}", errs) == "14", "named fact + expression + format")
    check(r.render("{t1.預算.sum}", errs) == "4,300", "sum with thousands separator")
    check(r.render("{t1[狀態!=正常].count}", errs) == "2", "!= filter + count")
    check(r.render("{t1[實際進度<70].count}", errs) == "3", "percent comparison without % sign")
    check(not errs, "no errors so far")
    r.render("{t1[代號=Z].實際進度}", errs)
    r.render("{t1.不存在}", errs)
    check(len(errs) == 2, "bad references are reported, not invented")


def test_build_and_patch():
    with open(EXAMPLE, encoding="utf-8") as f:
        spec = json.load(f)
    res = deck.build(spec, os.path.join(TMP, "out", "exec.pptx"))
    check(res["ok"], "example spec builds: %s" % res.get("errors"))
    texts = deck_texts(res["path"])
    check(len(texts) == 5, "cover + custom + 2 repeat pages + decision = 5 slides (got %d)" % len(texts))
    joined = "\n".join(texts)
    check("{" not in joined and "}" not in joined, "every {reference} resolved")
    check("供應鏈整合：延遲" in joined and "產線檢測：注意" in joined, "repeat pages expanded per project")
    check("14 個百分點" in joined, "computed gap rendered")
    check(len(res["annotations"]) == 2, "annotation table carried through")
    res2 = deck.patch(res["spec_path"], {"4": {"title": "需要決策：B 案外包（修改後）"}})
    check(res2["ok"], "patch_deck rebuilds")
    texts2 = deck_texts(res2["path"])
    check(len(texts2) == 5 and "修改後" in texts2[4] and texts2[1] == texts[1], "patch changed only the target slide")


def test_overflow_and_layouts():
    bullets = ["第 %d 項：這是一段很長的說明文字，用來測試自動縮字與自動拆成續頁的功能是否正常運作" % i for i in range(1, 26)]
    spec = {"title": "overflow", "audience": "engineer", "sources": [CSV], "slides": [
        {"title": "很多條列", "bullets": bullets},
        {"title": "三個區塊", "bullets": ["a", "b"], "kpis": [{"value": "{t1.count}", "label": "專案數"}],
         "chart": {"type": "column", "from": "t1", "category": "專案", "values": ["預算", "已用"], "sort": "-預算"}},
    ]}
    res = deck.build(spec, os.path.join(TMP, "out"))
    check(res["ok"], "overflow spec builds: %s" % res.get("errors"))
    check(len(res["mapping"][0]["slides"]) >= 2, "long text split into continuation slides")
    check(any("自訂排版" in w for w in res["warnings"]), "too many blocks -> auto custom layout")


def test_errors():
    bad = {"audience": "主管", "sources": [CSV], "slides": [
        {"title": "{t1[代號=Z].專案}"},
        {"layout": "custom", "regions": [{"type": "kpi", "content": [{"value": "1"}]}]},
        {"chart": {"type": "line", "from": "t1", "category": "專案", "values": ["不存在"]}},
    ]}
    res = deck.build(bad, os.path.join(TMP, "out"))
    check(not res["ok"] and len(res["errors"]) == 3, "unfixable problems are all returned at once (%d)" % len(res["errors"]))


def test_sketch_and_blueprint():
    paths = make_canvas(os.path.join(TMP, "canvas.pdf"), pages=3, requirements_page=True)
    check(len(load_sketch_pages(paths[0])) == 4, "multi-page PDF canvas readable")
    check([p["page"] for p in load_sketch_pages(paths[0], "2-3")] == [2, 3], "page ranges")
    r = tools.h_save_layout({"name": "月報-主管版", "slides": [{"regions": [
        {"id": "k", "type": "KPI", "x": 0, "y": 1.5, "w": 12, "h": 2},
        {"id": "b", "type": "重點", "x": 0, "y": 4, "w": 12, "h": 4}]}]})
    check(not r["isError"], "save_layout")
    spec = {"title": "bp", "audience": "manager", "sources": [CSV], "slides": [
        {"layout": "月報-主管版/1", "title": "用已存版型",
         "fill": {"k": [{"value": "{t1.count}", "label": "專案數"}], "b": ["第一點", "第二點"]}}]}
    res = deck.build(spec, os.path.join(TMP, "out"))
    check(res["ok"] and "第一點" in deck_texts(res["path"])[0], "blueprint layout + fill renders")


def _run_sizes(path):
    sizes = set()
    for slide in Presentation(path).slides:
        for shp in walk(slide.shapes):
            frames = []
            if shp.has_text_frame:
                frames.append(shp.text_frame)
            if getattr(shp, "has_table", False) and shp.has_table:
                frames.extend(c.text_frame for r in shp.table.rows for c in r.cells)
            for tf in frames:
                for p in tf.paragraphs:
                    sizes.update(r.font.size.pt for r in p.runs if r.font.size is not None)
    return sizes


def test_design_system():
    from pptx.enum.shapes import PP_PLACEHOLDER
    from report_engine import preview
    from report_engine.lint import lint_deck
    from report_engine.template import resolve_template_path
    from report_engine.theme import Theme, contrast, load_theme, status_key

    th = load_theme()
    check(th.primary == "1E4E8C" and th.font_ea == "Microsoft JhengHei", "theme.yaml loaded")
    check(Theme({"colors": {"primary": "#zzz"}}).primary == "1E4E8C", "invalid hex falls back to default")
    check(th.snap(15.3) == 14 and th.snap(100) == 54, "font sizes snap to the type scale")
    check(contrast(th.status_text("warn"), th.neutral(50)) >= 4.5, "amber status text darkened for contrast")
    check([status_key(x) for x in (u"延遲", "green", u"黃燈", "??")] == ["bad", "good", "warn", None], "status words")

    tpl = resolve_template_path()
    check(tpl.name == "_brand_16x9.pptx", "brand template generated when no company template")
    prs = Presentation(str(tpl))
    check((prs.slide_width, prs.slide_height) == (12192000, 6858000), "brand template is native 16:9")
    names = [L.name for L in prs.slide_masters[0].slide_layouts]
    check(names == [u"封面", u"標題及內容", u"章節", u"兩欄", u"僅標題", u"空白"], "brand layouts: %s" % names)
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme_xml = prs.slide_masters[0].part.part_related_by(RT.THEME).blob.decode("utf-8")
    check("Microsoft JhengHei" in theme_xml and "1E4E8C" in theme_xml, "theme XML carries brand font and colour")
    mtime = tpl.stat().st_mtime
    check(resolve_template_path().stat().st_mtime == mtime, "brand template not rebuilt when theme unchanged")
    theme_file = os.path.join(DESIGN, "theme.yaml")
    with open(theme_file, encoding="utf-8") as f:
        original = f.read()
    with open(theme_file, "w", encoding="utf-8") as f:
        f.write(original.replace('primary: "1E4E8C"', 'primary: "0F766E"'))
    rebuilt = Presentation(str(resolve_template_path()))
    check("0F766E" in rebuilt.slide_masters[0].part.part_related_by(RT.THEME).blob.decode("utf-8"),
          "editing theme.yaml rebuilds the brand template")
    with open(theme_file, "w", encoding="utf-8") as f:
        f.write(original)

    showcase = os.path.join(ROOT, "design_system", "examples", "showcase_diagrams.json")
    for spec_path, n in ((MANAGER, 5), (BENCH, 9), (showcase, 7)):
        with open(spec_path, encoding="utf-8") as f:
            spec = json.load(f)
        res = deck.build(spec, os.path.join(TMP, "ds_out", os.path.basename(spec_path) + ".pptx"),
                         base_dir=os.path.dirname(spec_path))
        check(res["ok"], "%s builds: %s" % (os.path.basename(spec_path), res.get("errors")))
        check(res["slide_count"] == n, "%s -> %d slides (got %d)" % (os.path.basename(spec_path), n, res["slide_count"]))
        warns = [i for i in res["lint"]["issues"] if i["level"] == "warn"]
        check(not warns, "%s passes the layout lint: %s" % (os.path.basename(spec_path), warns))
        off = _run_sizes(res["path"]) - set(th.steps)
        check(not off, "every font size comes from the type scale (off: %s)" % sorted(off))
        if spec_path == BENCH:
            bench_res = res
    res = bench_res
    texts = deck_texts(res["path"])
    check(u"延遲" in texts[2] and u"王雅婷" in texts[2], "status block rendered from table")
    check(u"落後專案：原因與對策" in texts[5], "section slide rendered")
    check(u"資料來源：" in texts[1] and texts[1].strip().endswith("2"), "footer source + slide number")
    pres = Presentation(res["path"])
    check(not any(s.is_placeholder and s.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER
                  for s in pres.slides[0].shapes), "no slide number on the cover")

    # lint catches overlap, overflow and poor contrast
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
    bad = Presentation(str(tpl))
    sl = bad.slides.add_slide(bad.slide_layouts[5])
    for i in range(2):
        tb = sl.shapes.add_textbox(Inches(1 + i * 0.5), Inches(2), Inches(4), Inches(1))
        r = tb.text_frame.paragraphs[0].add_run()
        r.text = u"重疊的文字方塊" * (12 if i else 1)
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor.from_string("DDDDDD")
    bad_path = os.path.join(TMP, "bad.pptx")
    bad.save(bad_path)
    msgs = " ".join(i["msg"] for i in lint_deck(bad_path)["issues"])
    check(u"重疊" in msgs and u"超出框" in msgs and u"對比" in msgs, "lint flags overlap / overflow / contrast")
    if os.name != "nt":
        imgs, err = preview.export_images(res["path"], None, 3)
        check(imgs is None and err, "preview degrades gracefully without PowerPoint")


def test_diagrams_svg_images():
    from PIL import Image, ImageDraw
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from report_engine.svg import path_points

    subs = path_points("M10 10 h 80 v 40 H 10 Z M 0 0 C 10 20 30 20 40 0 A 20 20 0 0 1 80 0 q 10 10 20 0")
    check(len(subs) == 2 and subs[0][1] and len(subs[0][0]) == 4 and subs[1][0][-1] == (100.0, 0.0),
          "SVG path parser: lines, curves, arcs, relative commands")
    shot = os.path.join(TMP, "shot.png")
    im = Image.new("RGB", (1400, 900), "white")
    d = ImageDraw.Draw(im)
    d.rectangle([40, 40, 1360, 860], fill="#F5F7FA")
    d.rectangle([40, 40, 1360, 100], fill="#1A67A3")
    for i in range(8):
        d.rectangle([300, 160 + i * 70, 1300, 200 + i * 70], fill="#FFFFFF" if i % 2 else "#E9EDF2")
    im.save(shot)
    spec = {"title": "diagrams", "audience": "manager", "slides": [
        {"layout": "custom", "title": "排程", "regions": [{"type": "甘特圖", "x": 0, "y": 1.3, "w": 12, "h": 6.7, "content": {
            "today": "2026/10/14", "tasks": [
                {"name": "設計", "start": "2026/08/01", "end": "2026/09/15", "progress": "100%", "group": "A"},
                {"name": "開發", "start": "2026-09-10", "end": "2026-12-20", "progress": 45, "status": "延遲", "group": "A"},
                {"name": "測試", "start": "2026/12/01", "end": "2027/01/10", "group": "B", "owner": "QA"}],
            "milestones": [{"name": "上線", "date": "2027/01/20"}]}}]},
        {"title": "時程", "timeline": {"today": "2026/10/14", "items": [
            {"date": "2026/04/15", "label": "啟動"}, {"date": "2026/06/01", "label": "設計", "note": "審查通過"},
            {"date": "2026/10/20", "label": "上線", "state": "current"}, {"date": "2026/12/15", "label": "結案"}]}},
        {"title": "流程", "process": {"steps": [{"title": "需求", "state": "done"}, {"title": "開發", "state": "current",
                                                                                   "status": "注意", "note": "人力不足"}, "測試"]}},
        {"title": "架構", "svg": "<svg viewBox='0 0 870 380'><rect x='10' y='10' width='200' height='80' rx='8' "
                                 "fill='primary'/><text x='110' y='55' text-anchor='middle' font-size='16' fill='white'>"
                                 "專案數 {t1.count}</text><line x1='210' y1='50' x2='400' y2='50' stroke='neutral-500' "
                                 "marker-end='url(#a)'/><circle cx='500' cy='50' r='40' fill='bad' opacity='0.5'/>"
                                 "<path d='M 600 20 l 100 0 l -50 80 z' fill='#E07A1F'/></svg>",
         "sources": []},
        {"layout": "custom", "title": "截圖", "regions": [{"type": "截圖", "x": 0, "y": 1.3, "w": 12, "h": 6.7, "content": {
            "path": shot, "style": "browser", "url": "openproject.example.com",
            "highlights": [{"x": 0.2, "y": 0.25, "w": 0.7, "h": 0.06, "note": "延遲的工作項"}],
            "zoom": {"x": 0.2, "y": 0.25, "w": 0.3, "h": 0.12}}}]},
        {"title": "照片", "images": {"items": [{"path": shot, "caption": "A"}, {"path": shot, "caption": "B", "style": "shadow"}]}},
    ], "sources": [CSV]}
    res = deck.build(spec, os.path.join(TMP, "diagrams", "d.pptx"))
    check(res["ok"], "diagram spec builds: %s" % res.get("errors"))
    warns = [i for i in res["lint"]["issues"] if i["level"] == "warn"]
    check(not warns, "diagram deck passes the layout lint: %s" % warns)
    pres = Presentation(res["path"])
    groups = [[s for s in sl.shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP] for sl in pres.slides]
    check([len(g) for g in groups] == [1, 1, 1, 1, 1, 2], "each drawing / image is one editable group")
    names = [g[0].name for g in groups]
    check(names == ["Gantt", "Timeline", "Process", "SVG", "Image", "Image"], "group names: %s" % names)

    def group_text(g):
        out = []
        for sh in g.shapes:
            if sh.has_text_frame:
                out.append(sh.text_frame.text)
        return " ".join(out)

    check(u"今天 10/14" in group_text(groups[0][0]) and "45%" in group_text(groups[0][0]), "gantt: today tag + progress")
    check(u"專案數 5" in group_text(groups[3][0]), "SVG text resolves {references}")
    pics = [sh for sh in groups[4][0].shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
    check(len(pics) == 2 and pics[0].crop_left > 0, "screenshot trimmed (native crop) + zoom inset")
    bad = deck.build({"audience": "manager", "slides": [
        {"svg": "<svg><rect></svg>"}, {"gantt": {"tasks": [{"name": "x", "start": "明天", "end": "2026/1/1"}]}}]},
        os.path.join(TMP, "diagrams"))
    check(not bad["ok"] and len(bad["errors"]) == 2, "broken SVG / bad dates are reported: %s" % bad.get("errors"))
    r = tools.h_view_image({"path": shot})
    check(not r["isError"] and r["content"][1]["type"] == "image" and "1400×900" in r["content"][0]["text"],
          "view_image returns a grid preview")


def test_workflow_assets_review():
    from PIL import Image
    from report_engine.review import review_deck

    adir = os.path.join(TMP, "assets", u"截圖")
    os.makedirs(adir)
    Image.new("RGB", (1600, 900), "#F5F7FA").save(os.path.join(adir, u"openproject_甘特圖_1014.png"))
    Image.new("RGB", (1200, 800), "#8FB3D9").save(os.path.join(adir, u"現場照片_AOI設備.jpg"))
    Image.new("RGB", (100, 100), "white").save(os.path.join(adir, "unrelated.png"))
    spec = {"title": "workflow", "audience": "manager", "sources": [CSV], "asset_dirs": [os.path.join(TMP, "assets")],
            "slides": [
                {"layout": "custom", "title": "排程現況", "regions": [
                    {"type": "截圖", "x": 0, "y": 1.3, "w": 8, "h": 6.7,
                     "content": {"need": u"OpenProject 甘特圖截圖", "search": ["gantt", u"甘特"], "style": "browser"}},
                    {"type": "bullets", "x": 8.4, "y": 1.6, "w": 3.6, "h": 6, "content": [u"專案數 {t1.count}"]}]},
                {"title": "現場", "image": {"need": u"AOI 設備現場照片"}},
                {"title": "資料錯誤", "bullets": ["{t1.不存在}"]}]}
    from report_engine.assets import collect
    text, imgs = collect(spec)
    check(u"openproject_甘特圖_1014.png" in text and u"現場照片_AOI設備.jpg" in text and "unrelated" not in text,
          "collect_assets finds candidate images by keyword")
    check(imgs and imgs[0]["type"] == "image", "collect_assets returns a numbered contact sheet")
    check(u"缺 1 項素材" in text and u"資料或寫法問題" in text, "plan table marks missing material + data errors")
    spec["slides"] = spec["slides"][:2]
    res = deck.build(spec, os.path.join(TMP, "workflow", "w.pptx"))
    check(res["ok"] and len(res["assets_needed"]) == 2, "deck builds with placeholders for missing images")
    check(u"待補圖片" in " ".join(deck_texts(res["path"])), "placeholder frame drawn on the slide")
    with open(BENCH, encoding="utf-8") as f:
        bench = json.load(f)
    bres = deck.build(bench, os.path.join(TMP, "workflow", "bench.pptx"), base_dir=os.path.dirname(BENCH))
    rv = review_deck(bres["path"])
    check(not [i for i in rv["issues"] if i["level"] == "warn"], "review_deck: benchmark deck is on-palette and aligned")
    odd = deck.build({"audience": "manager", "slides": [
        {"title": "x", "svg": "<svg viewBox='0 0 100 50'><rect width='100' height='50' fill='#00FF66'/></svg>"}]},
        os.path.join(TMP, "workflow", "odd.pptx"))
    msgs = " ".join(i["msg"] for i in review_deck(odd["path"])["issues"])
    check("00FF66" in msgs, "review_deck flags colours outside the brand palette")
    r = tools.call_tool("review_deck", {"path": bres["path"]})
    check(not r["isError"] and u"美術總檢" in r["content"][0]["text"], "review_deck tool")


def test_mcp_stdio():
    with open(EXAMPLE, encoding="utf-8") as f:
        spec = json.load(f)
    msgs = [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18", "capabilities": {}}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
         "params": {"name": "get_design_context", "arguments": {"audience": "主管", "sources": [CSV]}}},
        {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
         "params": {"name": "build_deck", "arguments": {"spec": spec, "out_path": os.path.join(TMP, "mcp")}}},
    ]
    data = "".join(json.dumps(m, ensure_ascii=False) + "\n" for m in msgs).encode("utf-8")
    proc = subprocess.run([sys.executable, os.path.join(ROOT, "mcp_server.py")], input=data,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=ROOT, env=os.environ.copy(), timeout=180)
    replies = [json.loads(line) for line in proc.stdout.decode("utf-8").splitlines() if line.strip()]
    check([r.get("id") for r in replies] == [1, 2, 3, 4], "one reply per request, none for the notification")
    check(replies[0]["result"]["serverInfo"]["name"] == "report-design", "initialize")
    check(len(replies[1]["result"]["tools"]) == 11, "tools/list returns 11 tools")
    ctx = replies[2]["result"]["content"][0]["text"]
    check("t1 ＝ sample_projects.csv" in ctx and "高階主管" in ctx, "context includes data and audience")
    check(u"視覺設計指引" in ctx and "preview_deck" in ctx, "context includes the visual design guide")
    check(not replies[3]["result"]["isError"] and "已產出" in replies[3]["result"]["content"][0]["text"],
          "build_deck over MCP")


if __name__ == "__main__":
    try:
        for fn in (test_facts, test_build_and_patch, test_overflow_and_layouts, test_errors,
                   test_sketch_and_blueprint, test_design_system, test_diagrams_svg_images,
                   test_workflow_assets_review, test_mcp_stdio):
            print("==", fn.__name__)
            fn()
        print("\nALL PASSED  (outputs in %s)" % TMP)
    except Exception:
        print("\nFAILED  (outputs in %s)" % TMP)
        raise
