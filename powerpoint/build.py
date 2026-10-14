#!/usr/bin/env python3
"""Build the University of Nottingham PowerPoint template with OfficeCLI.

Design note: OfficeCLI keeps an in-memory "resident" per file. Spawning a
*separate* officecli process to read the file while the resident holds it can
shadow/reload and silently drop unflushed edits. So this script drives the
whole build inside ONE shell session — every mutating call plus the final
`close` happen in a single `bash -c` chain, and after `close` we never invoke
officecli on this file again (verification/render use zipfile or the
already-flushed bytes on disk).

Reads spec.json (single source of truth for colours/fonts/geometry).

Run:  python3 build.py [--step all|plan|validate|potx|render]
"""
import argparse
import json
import os
import shlex
import subprocess
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SPEC_PATH = os.path.join(HERE, "spec.json")
PPTX = os.path.join(HERE, "UoN-Presentation-Template.pptx")
POTX = os.path.join(HERE, "UoN-Presentation-Template.potx")
PREVIEW = os.path.join(HERE, "preview")
ASSETS = os.path.join(HERE, "assets")

PML_PRESENTATION_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
PML_TEMPLATE_MAIN = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"


def sh(cmd, timeout=600):
    p = subprocess.run(["bash", "-c", cmd], capture_output=True, text=True, timeout=timeout)
    return p.returncode, p.stdout


def q(s):
    return shlex.quote(str(s))


def pt(v):
    return f"{v}pt"


# --------------------------------------------------------------------- builders
def set_cmd(path, **kw):
    props = []
    for k, v in kw.items():
        if v is None:
            continue
        if isinstance(v, bool):
            v = "true" if v else "false"
        props.append(f"--prop={k}={v}")
    return f"officecli set {q(PPTX)} {q(path)} " + " ".join(q(p) for p in props) + " --json"


def shape_cmd(parent, *, text=None, x, y, w, h, fill=None, color=None,
              size=None, bold=None, align=None, geometry=None, font=None,
              valign=None, image=None, italic=None):
    props = [f"--prop=x={pt(x)}", f"--prop=y={pt(y)}",
             f"--prop=width={pt(w)}", f"--prop=height={pt(h)}"]
    if text is not None:
        props.append(f"--prop=text={text}")
    if fill:
        props.append(f"--prop=fill={fill}")
    if color:
        props.append(f"--prop=color={color}")
    if size:
        props.append(f"--prop=size={size}")
    if bold is not None:
        props.append(f"--prop=bold={'true' if bold else 'false'}")
    if italic is not None:
        props.append(f"--prop=italic={'true' if italic else 'false'}")
    if align:
        props.append(f"--prop=align={align}")
    if valign:
        props.append(f"--prop=valign={valign}")
    if geometry:
        props.append(f"--prop=geometry={geometry}")
    if font:
        props.append(f"--prop=font={font}")
    if image:
        props.append(f"--prop=image={image}")
    return f"officecli add {q(PPTX)} {q(parent)} --type shape " + " ".join(q(p) for p in props) + " --json"


def ph_cmd(slide, *, ph_type, text, x=None, y=None, w=None, h=None,
           size=None, bold=None, align=None, color=None, font=None, valign=None,
           idx=None):
    props = [f"--prop=phType={ph_type}"]
    if idx is not None:
        props.append(f"--prop=idx={idx}")
    if text is not None:
        props.append(f"--prop=text={text}")
    if x is not None:
        props.append(f"--prop=x={pt(x)}")
    if y is not None:
        props.append(f"--prop=y={pt(y)}")
    if w is not None:
        props.append(f"--prop=width={pt(w)}")
    if h is not None:
        props.append(f"--prop=height={pt(h)}")
    if size:
        props.append(f"--prop=size={size}")
    if bold is not None:
        props.append(f"--prop=bold={'true' if bold else 'false'}")
    if align:
        props.append(f"--prop=align={align}")
    if valign:
        props.append(f"--prop=valign={valign}")
    if color:
        props.append(f"--prop=color={color}")
    if font:
        props.append(f"--prop=font={font}")
    return f"officecli add {q(PPTX)} {q(slide)} --type placeholder " + " ".join(q(p) for p in props) + " --json"


def slide_cmd(layout, *, title=None, text=None, name=None, subtitle=None, background=None):
    props = [f"--prop=layout={layout}"]
    if title is not None:
        props.append(f"--prop=title={title}")
    if subtitle is not None:
        props.append(f"--prop=subtitle={subtitle}")
    if text is not None:
        props.append(f"--prop=text={text}")
    if name:
        props.append(f"--prop=name={name}")
    if background:
        props.append(f"--prop=background={background}")
    return f"officecli add {q(PPTX)} / --type slide " + " ".join(q(p) for p in props) + " --json"


def chart_cmd(slide, *, chart_type, x, y, w, h, categories, series,
              title=None, series_colors=None):
    # scatter/bubble map `categories` onto the X axis (per officecli chart docs).
    # String labels (Q1..Q4) can collapse to x=0 in the renderer — supply numeric
    # X values for those chart types instead of the display labels.
    if chart_type in ("scatter", "bubble"):
        cats = ",".join(str(i + 1) for i in range(len(categories)))
    else:
        cats = ",".join(str(c) for c in categories)
    props = [f"--prop=chartType={chart_type}", f"--prop=x={pt(x)}", f"--prop=y={pt(y)}",
             f"--prop=width={pt(w)}", f"--prop=height={pt(h)}", f"--prop=categories={cats}"]
    if title:
        props.append(f"--prop=title={title}")
    for i, (nm, vals) in enumerate(series, start=1):
        props.append(f"--prop=series{i}.name={nm}")
        props.append(f"--prop=series{i}.values={','.join(str(v) for v in vals)}")
    if series_colors:
        for i, c in enumerate(series_colors, start=1):
            props.append(f"--prop=series{i}.color={c}")
        # some chart types (waterfall/funnel/histogram/pie/doughnut) read chart-level
        # `colors` instead of per-series. NOTE: 'pareto' removed from this tuple —
        # officecli's chartEx pareto renderer shuffles categories and ignores these;
        # we build pareto as a native column+line combo instead (see below).
        if chart_type in ("waterfall", "funnel", "histogram", "pie", "doughnut"):
            props.append(f"--prop=colors={','.join(series_colors)}")

    # Native Pareto: column (bars, primary axis) + line (cumulative %, secondary axis).
    # officecli supports combos via combotypes + secondaryaxis props — this renders a
    # genuine Pareto (descending bars + cumulative line) with correct category order,
    # unlike the chartEx `pareto` chartType which has data-shuffling bugs.
    if chart_type == "pareto":
        props = [p for p in props if not p.startswith("--prop=chartType=")]
        props.insert(0, "--prop=chartType=column")
        props.append(f"--prop=combotypes=column,line")
        props.append(f"--prop=combosplit=1")          # split series 1 (bars) / 2 (line)
        props.append(f"--prop=secondaryaxis=2")        # series 2 on secondary axis
        props.append(f"--prop=axismax=100")            # secondary axis 0-100%
        # ensure line series colour stays brand primary for the cumulative line
        if len(series_colors) >= 2:
            props.append(f"--prop=series2.color={series_colors[1]}")
    return f"officecli add {q(PPTX)} {q(slide)} --type chart " + " ".join(q(p) for p in props) + " --json"


def table_cmd(slide, *, x, y, w, h, rows, header_fill=None, body_fill=None):
    """rows: list of lists. officecli table `data` prop is CSV-ish:
    'H1,H2;R1C1,R1C2' (quote-aware). `rows` prop is an int count only."""
    def esc(v):
        s = str(v)
        # quote-aware: wrap in double quotes if cell contains comma/quote/semicolon
        if any(ch in s for ch in (',', ';', '"')):
            s = '"' + s.replace('"', '""') + '"'
        return s
    data_csv = ";".join(",".join(esc(v) for v in row) for row in rows)
    props = [f"--prop=x={pt(x)}", f"--prop=y={pt(y)}",
             f"--prop=width={pt(w)}", f"--prop=height={pt(h)}",
             f"--prop=data={data_csv}"]
    if header_fill:
        props.append(f"--prop=headerFill={header_fill}")
    if body_fill:
        props.append(f"--prop=bodyFill={body_fill}")
    return f"officecli add {q(PPTX)} {q(slide)} --type table " + " ".join(q(p) for p in props) + " --json"


# ------------------------------------------------------------------ build plan
def plan_commands():
    spec = json.load(open(SPEC_PATH, encoding="utf-8"))
    t, g, C = spec["theme"], spec["geometry_pt"], spec["colours"]
    primary, white, sky = C["primary"]["hex"], C["neutral"]["white"], C["supporting"]["malaysia_sky_blue"]
    tint5, blue60 = C["tints"]["blue5"], C["tints"]["blue60"]
    cmds = []

    theme_props = [
        f"--prop=headingFont={t['headingFont']}", f"--prop=bodyFont={t['bodyFont']}",
        f"--prop=headingFont.ea={t['headingFont_ea']}", f"--prop=bodyFont.ea={t['bodyFont_ea']}",
        f"--prop=dk1=#{t['dk1']}", f"--prop=lt1=#{t['lt1']}",
        f"--prop=dk2=#{t['dk2']}", f"--prop=lt2=#{t['lt2']}",
        f"--prop=accent1=#{t['accent1']}", f"--prop=accent2=#{t['accent2']}",
        f"--prop=accent3=#{t['accent3']}", f"--prop=accent4=#{t['accent4']}",
        f"--prop=accent5=#{t['accent5']}", f"--prop=accent6=#{t['accent6']}",
        f"--prop=hyperlink=#{t['hyperlink']}", f"--prop=followedhyperlink=#{t['followedhyperlink']}",
    ]
    cmds.append(f"officecli set {q(PPTX)} /theme " + " ".join(q(p) for p in theme_props) + " --json")

    logo = os.path.join(ASSETS, "logo-white.png")
    layout_specs = {
        "/slidelayout[1]": {"name": "Blank", "header": True, "progress": True, "page_no": True},
        "/slidelayout[2]": {"name": "Cover", "header": False, "progress": False, "page_no": False, "jumbo": True},
        "/slidelayout[3]": {"name": "Title and Content", "header": True, "progress": True, "page_no": True},
        "/slidelayout[4]": {"name": "Two Content", "header": True, "progress": True, "page_no": True},
        "/slidelayout[5]": {"name": "Title Only", "header": True, "progress": True, "page_no": True},
    }
    for path, cfg in layout_specs.items():
        cmds.append(set_cmd(path, name=cfg["name"]))
        if cfg.get("jumbo"):
            # Cover: large dark band at bottom (avoid covering title/subtitle on left)
            jb = g.get("cover_jumbo", {"x": 0, "y": 340, "w": 960, "h": 200})
            cmds.append(shape_cmd(path, text="__JUMBO__", x=jb["x"], y=jb["y"],
                                  w=jb["w"], h=jb["h"], fill=primary))
        if cfg["header"]:
            cmds.append(shape_cmd(path, text="__HEADER__", x=g["header_band"]["x"],
                                  y=g["header_band"]["y"], w=g["header_band"]["w"],
                                  h=g["header_band"]["h"], fill=primary, color=primary))
            # Wordmark on the header band (right-aligned). NOTE: OfficeCLI cannot
            # persist an image fill on a *layout* (blipFill is dropped on close;
            # `add --type picture` also refuses non-slide parents), so we fall back
            # to a white text wordmark which renders identically in PowerPoint.
            lb = g.get("logo_on_band", {"x": 560, "y": 6, "w": 384, "h": 32})
            cmds.append(shape_cmd(path, text="University of Nottingham",
                                  x=lb["x"], y=lb["y"], w=lb["w"], h=lb["h"],
                                  size=14, color="FFFFFF", align="right",
                                  valign="middle", font=t["bodyFont"]))
        if cfg["progress"]:
            cmds.append(shape_cmd(path, text="__PROGTRACK__", x=g["progress_track"]["x"],
                                  y=g["progress_track"]["y"], w=g["progress_track"]["w"],
                                  h=g["progress_track"]["h"], fill=tint5, color=tint5))
            cmds.append(shape_cmd(path, text="__PROGBAR__", x=g["progress_bar"]["x"],
                                  y=g["progress_bar"]["y"], w=g["progress_bar"]["w"],
                                  h=g["progress_bar"]["h"], fill=sky, color=sky))
        # page_no handled by raw-set <a:fld type="slidenum"/> injected after layouts (below)

    # slide-number field on every layout flagged page_no=True (auto-numbered by PowerPoint)
    fld_xml = (
        '<p:sp><p:nvSpPr><p:cNvPr id="9001" name="SlideNumber"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:off x="11430000" y="6449560"/><a:ext cx="635000" cy="228600"/></a:xfrm>'
        '<a:prstGeom prst="rect" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:avLst/></a:prstGeom></p:spPr>'
        '<p:txBody><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" rtlCol="0"/>'
        '<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:lvl1pPr algn="r"/></a:lstStyle>'
        '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:pPr algn="r"/>'
        '<a:fld id="{F1E1D0C0-B0B1-4A01-9E01-000000000001}" type="slidenum"><a:rPr lang="en-US" sz="1100" dirty="0"/>'
        '<a:t>1</a:t></a:fld></a:p></p:txBody></p:sp>'
    )
    for lp, cfg in layout_specs.items():
        if cfg.get("page_no"):
            part = lp.replace("/slidelayout[", "/slideLayout[")
            cmds.append(f"officecli raw-set {q(PPTX)} {q(part)} "
                        f"--xpath \"//p:spTree\" --action append --xml {q(fld_xml)} --json")

    # NOTE: layout4 (Two Content) already carries TWO body placeholders (idx=1 left, idx=2 right)
    # out of the box — do NOT add another. We only shrink the LEFT one to half-width so the
    # right column has room. layout4 shape order: [1]=Title, [2]=left body(idx1), [3]=right body(idx2).
    cmds.append(set_cmd("/slidelayout[4]/shape[2]", width=pt(430), height=pt(330)))
    cmds.append(set_cmd("/slidelayout[4]/shape[3]", x=pt(504), width=pt(430), height=pt(330)))

    # NOTE: bullet list style injection moved to post-build fix_body_bullets() —
    # officecli raw-set `--action append` into a self-closing <a:lstStyle/> is unreliable
    # (lvl1pPr lands outside the tag / command fails). Deterministic zip-layer fix instead.

    # move placeholders down below the header band so title/body don't collide.
    # OfficeCLI `set` only accepts numeric shape indices on layouts (no [@phType=...]).
    # Layout shape order (verified): layout3: [1]=Title [2]=Content; layout4: [1]=Title [2][3]=body;
    # layout5: [1]=Title; layout2: [1]=ctrTitle [2]=subTitle.
    cmds.append(set_cmd("/slidelayout[3]/shape[1]", y=pt(56), height=pt(64)))
    cmds.append(set_cmd("/slidelayout[3]/shape[2]", y=pt(168), height=pt(330)))
    cmds.append(set_cmd("/slidelayout[4]/shape[1]", y=pt(56), height=pt(64)))
    cmds.append(set_cmd("/slidelayout[5]/shape[1]", y=pt(56), height=pt(64)))
    cmds.append(set_cmd("/slidelayout[2]/shape[1]", y=pt(120), height=pt(90)))
    cmds.append(set_cmd("/slidelayout[2]/shape[2]", y=pt(220), height=pt(70)))
    return cmds


def sample_commands():
    spec = json.load(open(SPEC_PATH, encoding="utf-8"))
    C = spec["colours"]
    primary = C["primary"]["hex"]
    stone = C["neutral"]["portland_stone"]
    sky = C["supporting"]["malaysia_sky_blue"]
    white = C["neutral"]["white"]
    blue60 = C["tints"]["blue60"]
    items = []

    items.append(slide_cmd("Cover", title="Nottingham Presentation Template", name="Cover"))
    items.append(("PH", "Cover", "subtitle", "UoN brand-compliant template", white))
    items.append(slide_cmd("Title Only", background=primary, name="Section"))
    items.append(("TEXTBOX", "Section", "Section Divider", 80, 240, 800, 80, 44, white, True, "left", "middle"))
    items.append(slide_cmd("Title and Content", title="Title and Content",
                           text=" First key point\n Second key point\n Third key point",
                           name="Content"))
    items.append(slide_cmd("Two Content", title="Two Content",
                           text=" Left column point A\n Left column point B", name="TwoContent"))
    items.append(("PH_BODY2", "TwoContent", " Right column point A\n Right column point B"))
    items.append(slide_cmd("Two Content", title="Comparison",
                           text=" Before\n After", name="Comparison"))
    items.append(("PH_BODY2", "Comparison", " After (or contrast point)\n Another point"))
    items.append(slide_cmd("Title Only", title="Title Only", name="TitleOnly"))
    items.append(slide_cmd("Blank", title="Blank Layout", name="Blank"))

    cats = ["Q1", "Q2", "Q3", "Q4"]
    series = [("2025", [12, 19, 15, 22]), ("2026", [15, 22, 18, 27])]
    colors = [sky, primary]
    for ct in ["bar", "column", "line", "area", "pie", "doughnut", "scatter",
               "radar", "stackedColumn", "percentStackedColumn", "waterfall",
               "funnel", "pareto", "histogram"]:
        items.append(slide_cmd("Title Only", title=f"Chart: {ct}", name=f"chart-{ct}"))
        items.append(("CHART", ct, cats, series, colors))

    items.append(slide_cmd("Title Only", title="Table", name="Table"))
    items.append(("TABLE", [["Metric", "Q1", "Q2", "Q3"],
                            ["Revenue", "120", "150", "180"],
                            ["Growth", "8%", "11%", "14%"]]))

    items.append(slide_cmd("Blank", background=primary, name="Statement"))
    items.append(("TEXTBOX", "Statement", "A single, bold statement goes here.", 80, 220, 800, 100, 40, white, True, "center", "middle"))
    items.append(slide_cmd("Blank", background=stone, name="Quote"))
    items.append(("TEXTBOX", "Quote", "\"Quotes make a point land.\"", 80, 200, 800, 90, 32, primary, False, "center", "middle"))
    items.append(("TEXTBOX", "Quote", "— Author", 80, 300, 800, 40, 20, blue60, False, "right", "middle"))
    items.append(slide_cmd("Blank", background=primary, name="Closing"))
    items.append(("TEXTBOX", "Closing", "Thank you", 80, 230, 800, 90, 44, white, True, "center", "middle"))
    return items


def build_script():
    base = plan_commands()
    samp = sample_commands()
    lines = ["set -e", f"officecli open {q(PPTX)} >/dev/null 2>&1 || true"]

    def emit(cmd):
        lines.append(cmd + ' >/dev/null 2>&1 || echo "FAILED: ' + cmd[:70] + '"')

    for c in base:
        emit(c)

    slide_idx = 0
    for item in samp:
        if isinstance(item, str):
            if "--type slide" in item:
                slide_idx += 1
            emit(item)
        elif item[0] == "CHART":
            _, ct, cats, series, colors = item
            emit(chart_cmd(f"/slide[{slide_idx}]", chart_type=ct, x=60, y=140, w=840, h=330,
                           categories=cats, series=series, title=ct, series_colors=colors))
        elif item[0] == "TABLE":
            _, rows = item
            emit(table_cmd(f"/slide[{slide_idx}]", x=60, y=150, w=840, h=240,
                           rows=rows, header_fill=C_PRIMARY, body_fill=C_STONE))
            # header row cells: headerFill sets FILL but not text colour — without an
            # explicit white the dark-blue header renders invisible text.
            n_cols = len(rows[0])
            for c in range(1, n_cols + 1):
                emit(f"officecli set {q(PPTX)} {q(f'/slide[{slide_idx}]/table[1]/tr[1]/tc[{c}]')} "
                     f"--prop=color=FFFFFF --prop=bold=true --json")
        elif item[0] == "PH_BODY2":
            _, name, txt = item
            emit(ph_cmd(f"/slide[{slide_idx}]", ph_type="body", text=txt, idx=2,
                        x=504, y=168, w=430, h=330, size=20, color=C_PRIMARY, font="Arial"))
        elif item[0] == "PH":
            _, lay_name, ph_type, txt = item[:4]
            col = item[4] if len(item) > 4 else C_STONE
            emit(ph_cmd(f"/slide[{slide_idx}]", ph_type=ph_type, text=txt,
                        x=54, y=306, w=504, h=60, size=20, color=col))
        elif item[0] == "TEXTBOX":
            _, name, txt, x, y, w, h, size, color, bold, align, valign = item
            emit(shape_cmd(f"/slide[{slide_idx}]", text=txt, x=x, y=y, w=w, h=h,
                           size=size, color=color, bold=bold, align=align, valign=valign,
                           font="Arial"))

    lines.append(f"officecli close {q(PPTX)} >/dev/null 2>&1 || true")
    return "\n".join(lines) + "\n"


C_PRIMARY = "10263B"
C_STONE = "FAF6EF"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--step", default="all")
    args = ap.parse_args()

    if args.step == "potx":
        make_potx()
        return
    if args.step == "validate":
        rc, out = sh(f"officecli validate {q(PPTX)}", timeout=120)
        print("[validate]", out.strip() or f"rc={rc}")
        return
    if args.step == "render":
        os.makedirs(PREVIEW, exist_ok=True)
        rc, out = sh(f"officecli view {q(PPTX)} screenshot --grid auto -o {q(os.path.join(PREVIEW, 'overview.png'))}", timeout=300)
        print("[render]", out.strip() or f"rc={rc}")
        return

    script = build_script()
    if args.step == "plan":
        print(script)
        return

    tmp = os.path.join(HERE, "_build.sh")
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(script)
    rc, out = sh(f"bash {q(tmp)}", timeout=900)
    print("[build rc]", rc)
    if out.strip():
        print(out[-3000:])

    if args.step in ("all", "validate"):
        rc, out = sh(f"officecli validate {q(PPTX)}", timeout=120)
        print("[validate]", out.strip() or f"rc={rc}")
    if args.step in ("all", "potx"):
        fix_waterfall_colours()
        fix_body_bullets()
        make_potx()
    if args.step in ("all", "render"):
        os.makedirs(PREVIEW, exist_ok=True)
        rc, out = sh(f"officecli view {q(PPTX)} screenshot --grid auto -o {q(os.path.join(PREVIEW, 'overview.png'))}", timeout=300)
        print("[render]", out.strip() or f"rc={rc}")


def fix_waterfall_colours():
    """officecli generates waterfall charts with Office default colours (4472C4/2E75B6/FF0000)
    regardless of seriesN.color props — remap to brand palette post-build."""
    import re as _re
    z = zipfile.ZipFile(PPTX)
    wf_part = None
    for n in sorted(z.namelist()):
        if "charts/chart" in n and n.endswith(".xml"):
            c = z.read(n).decode()
            if "Increase" in c and "Decrease" in c:
                wf_part = n
                break
    if not wf_part:
        z.close()
        return
    c = z.read(wf_part).decode()
    z.close()
    c2 = (c.replace('val="4472C4"', 'val="009BC1"')
            .replace('val="2E75B6"', 'val="10263B"')
            .replace('val="FF0000"', 'val="B91C2E"'))
    if c2 == c:
        print("[waterfall] already brand-coloured")
        return
    tmp = PPTX + ".tmp"
    with zipfile.ZipFile(PPTX) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == wf_part:
                data = c2.encode()
            zout.writestr(item, data)
    os.replace(tmp, PPTX)
    print(f"[waterfall] {wf_part} colours -> brand")


def fix_body_bullets():
    """OfficeCLI raw-set `--action append` into a self-closing <a:lstStyle/> is unreliable
    (the lvl1pPr lands outside the tag, or the command fails). Do it deterministically at
    the zip layer: for each layout's body placeholder, expand the self-closing lstStyle and
    inject the bullet lvl1pPr inside it."""
    import re as _re
    bullet = ('<a:lvl1pPr marL="285750" indent="-285750" algn="l">'
              '<a:spcBef><a:spcPts val="600"/></a:spcBef>'
              '<a:buFont typeface="Arial"/><a:buChar char="•"/></a:lvl1pPr>')
    # match a body ph's self-closing <a:lstStyle .../>  → expand + inject
    # (simplified non-greedy pattern — the original nested-lookahead version had an
    #  unbalanced parenthesis and crashed the build)
    body_pat = _re.compile(
        r'(<p:sp>.*?<p:ph[^>]*type="body".*?<a:lstStyle[^>]*)/>',
        _re.S)
    rewrite = {}
    z = zipfile.ZipFile(PPTX)
    for i in range(1, 6):
        part = f"ppt/slideLayouts/slideLayout{i}.xml"
        if part not in z.namelist():
            continue
        lay = z.read(part).decode()
        # only inject if this layout's body lstStyle has no lvl1pPr yet (idempotent)
        def repl(m):
            return m.group(1) + ">" + bullet + "</a:lstStyle>"
        new = body_pat.sub(repl, lay)
        if new != lay:
            rewrite[part] = new
    z.close()
    if not rewrite:
        print("[bullets] no layout needed bullet injection")
        return
    tmp = PPTX + ".tmp"
    with zipfile.ZipFile(PPTX) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in rewrite:
                data = rewrite[item.filename].encode()
            zout.writestr(item, data)
    os.replace(tmp, PPTX)
    print(f"[bullets] injected into layouts: {sorted(rewrite)}")


def make_potx():
    if os.path.exists(POTX):
        os.remove(POTX)
    zin = zipfile.ZipFile(PPTX)
    with zipfile.ZipFile(POTX, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(PML_PRESENTATION_MAIN.encode(), PML_TEMPLATE_MAIN.encode())
            zout.writestr(item, data)
    zin.close()
    z = zipfile.ZipFile(POTX)
    ct = z.read("[Content_Types].xml").decode()
    assert PML_TEMPLATE_MAIN in ct and PML_PRESENTATION_MAIN not in ct, "content type swap failed"
    print(f"[potx] {os.path.basename(POTX)} ({os.path.getsize(POTX)} bytes) -> template.main+xml")


if __name__ == "__main__":
    main()
