"""Build an editable PPTX from per-figure specs + detected geometry.

fig2ppt v1.0.0 (2026-09-26) - packaged from the original _build pipeline.
Usage: python build_pptx.py [spec_a.json spec_b.json ... out.pptx]
       No args: builds every spec_*.json in the current dir -> out.pptx
"""
import io
import json
import glob
import os
import sys
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_PATTERN, MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SLIDE_W_EMU = int(13.3333 * 914400)
SLIDE_H_EMU = int(7.5 * 914400)
FONT = "Times New Roman"


def emu(px, scale):
    return Emu(int(round(px * scale)))


def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def darker(h, f=0.62):
    c = [int(h.lstrip('#')[i:i + 2], 16) for i in (0, 2, 4)]
    return "#%02X%02X%02X" % tuple(max(0, int(v * f)) for v in c)


SHAPE_MAP = {
    "rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "rect_sharp": MSO_SHAPE.RECTANGLE,
    "ellipse_ish": MSO_SHAPE.OVAL,
    "ellipse": MSO_SHAPE.OVAL,
    "diamond": MSO_SHAPE.DIAMOND,
    "triangleR": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "triangleL": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "triangleU": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "hexagon": MSO_SHAPE.HEXAGON,
    "hexagon_v": MSO_SHAPE.HEXAGON,
    "parallelogram": MSO_SHAPE.PARALLELOGRAM,
    "trapezoid": MSO_SHAPE.TRAPEZOID,
    "pentagon": MSO_SHAPE.PENTAGON,
    "cylinder": MSO_SHAPE.CAN,
}


def add_arrowhead(line_elem, kind="triangle", size="med"):
    ln = line_elem._get_or_add_ln()
    for tag in ("a:tailEnd",):
        e = ln.find(qn(tag))
        if e is None:
            e = ln.makeelement(qn(tag), {})
            ln.append(e)
        e.set("type", kind)
        e.set("w", size)
        e.set("len", size)


def classify_size(w_px, h_px, text, scale, override=None):
    if override:
        return Pt(override)
    pt_per_px = scale * 72.0
    h_pt = h_px * pt_per_px
    w_pt = w_px * pt_per_px
    if not text:
        return Pt(8)
    lines = text.split("\n")
    n = len(lines)
    maxlen = max(len(l) for l in lines)
    by_h = h_pt / (n * 1.28)
    by_w = w_pt / (0.52 * max(1, maxlen)) * 1.05
    return Pt(max(5.5, min(20.0, by_h, by_w)))


def set_text(shape, text, size, align="c", color="#000000", bold=False, font=FONT):
    tf = shape.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Emu(9525)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"c": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT}[align]
        r = p.add_run()
        r.text = ln
        r.font.size = size
        r.font.name = font
        r.font.color.rgb = rgb(color)
        r.font.bold = bold


def anchor_pt(box, side):
    x, y, w, h = box["x"], box["y"], box["w"], box["h"]
    if side == "l":
        return (x, y + h / 2)
    if side == "r":
        return (x + w, y + h / 2)
    if side == "t":
        return (x + w / 2, y)
    if side == "b":
        return (x + w / 2, y + h)
    return (x + w / 2, y + h / 2)


def parse_anchor(token, by_id):
    """token: '12:r' (side l/r/t/b/c, default c) or '#x,y'"""
    if token.startswith("#"):
        xs, ys = token[1:].split(",")
        return (float(xs), float(ys))
    parts = token.split(":")
    side = parts[1] if len(parts) > 1 and parts[1] else "c"
    return anchor_pt(by_id[int(parts[0])], side)


def add_polyline(slide, pts, to_emu, color, width_pt=1.0, arrow=True, dash=False):
    conv = [to_emu(x, y) for x, y in pts]
    dedup = []
    for p in conv:
        if not dedup or (p[0] != dedup[-1][0] or p[1] != dedup[-1][1]):
            dedup.append(p)
    if len(dedup) < 2:
        return None
    fb = slide.shapes.build_freeform(dedup[0][0], dedup[0][1], scale=1.0)
    fb.add_line_segments(dedup[1:], close=False)
    shp = fb.convert_to_shape()
    shp.fill.background()
    shp.shadow.inherit = False
    ln = shp.line
    ln.color.rgb = rgb(color)
    ln.width = Pt(width_pt)
    if dash:
        ln.dash_style = MSO_LINE_DASH_STYLE.DASH
    if arrow:
        add_arrowhead(ln)
    return shp


def elbow_points(p1, p2, mode="hv"):
    (x1, y1), (x2, y2) = p1, p2
    if mode == "hvh":
        xm = (x1 + x2) / 2
        return [p1, (xm, y1), (xm, y2), p2]
    if mode == "vhv":
        ym = (y1 + y2) / 2
        return [p1, (x1, ym), (x2, ym), p2]
    if abs(x2 - x1) >= abs(y2 - y1):
        return elbow_points(p1, p2, "hvh")
    return elbow_points(p1, p2, "vhv")


def build(spec_paths, out_path):
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W_EMU)
    prs.slide_height = Emu(SLIDE_H_EMU)
    blank = prs.slide_layouts[6]
    report = []
    for sp_path in spec_paths:
        with open(sp_path, encoding="utf-8") as f:
            spec = json.load(f)
        sp_name = os.path.basename(sp_path)
        img = spec["image"]
        geo_path = os.path.join(os.path.dirname(sp_path), os.path.splitext(img)[0] + ".json")
        with open(geo_path, encoding="utf-8") as f:
            geo = json.load(f)
        by_id = {b["id"]: b for b in geo["boxes"]}
        IW, IH = geo["W"], geo["H"]
        slide = prs.slides.add_slide(blank)
        scale = min(SLIDE_W_EMU * 0.985 / IW, SLIDE_H_EMU * 0.965 / IH) / 914400.0
        off_x = (13.3333 - IW * scale) / 2
        off_y = (7.5 - IH * scale) / 2
        font = spec.get("font", FONT)

        known = {str(b["id"]) for b in geo["boxes"]}
        for k in spec.get("overrides", {}):
            if k not in known:
                print(f"WARN {sp_name}: override id {k} 不在 detect 结果中")
        for m in spec.get("merges", []):
            miss = [i for i in m["ids"] if i not in by_id]
            if miss:
                print(f"WARN {sp_name}: merge ids {miss} 不在 detect 结果中")

        def to_emu(px, py, _s=scale, _ox=off_x, _oy=off_y):
            return (Emu(int((px * _s + _ox) * 914400)), Emu(int((py * _s + _oy) * 914400)))

        def X(px):
            return to_emu(px, 0)[0]

        def Y(px):
            return to_emu(0, px)[1]

        drops = set(spec.get("drops", []))
        keep_only = spec.get("keep_only", False)
        if keep_only:
            keep_ids = set(spec.get("overrides", {}).keys()) | {str(i) for m in spec.get("merges", []) for i in m["ids"]}
            boxes = [dict(b) for b in geo["boxes"] if str(b["id"]) in keep_ids]
        else:
            boxes = [dict(b) for b in geo["boxes"] if b["id"] not in drops]
        for m in spec.get("merges", []):
            ids = m["ids"]
            parts = [by_id[i] for i in ids if i in by_id]
            if not parts:
                continue
            x0 = min(p["x"] for p in parts); y0 = min(p["y"] for p in parts)
            x1 = max(p["x"] + p["w"] for p in parts); y1 = max(p["y"] + p["h"] for p in parts)
            boxes = [b for b in boxes if b["id"] not in ids]
            boxes.append(dict(parts[0], id=-1, x=x0, y=y0, w=x1 - x0, h=y1 - y0,
                              text=m.get("text", ""), shape=m.get("shape", "rect"),
                              fill=m.get("fill", parts[0]["fill"]),
                              stroke=m.get("stroke", parts[0]["stroke"])))
        for b in boxes:
            ov = spec.get("overrides", {}).get(str(b["id"]), {})
            for k in ("shape", "fill", "stroke", "size", "dash", "no_line", "text", "text_color"):
                if k in ov:
                    b[k] = ov[k]
            b.setdefault("text", "")
        for b in spec.get("add", []):
            r = b["rect"]
            boxes.append(dict(id=-1, x=r[0], y=r[1], w=r[2], h=r[3],
                              text=b.get("text", ""), fill=b.get("fill", "#FFFFFF"),
                              stroke=b.get("stroke", "#000000"), shape=b.get("shape", "rect"),
                              size=b.get("size"), dash=b.get("dash", False),
                              no_line=b.get("no_line", False), no_fill=b.get("no_fill", False),
                              icon=b.get("icon"), rot=b.get("rot", 0), crot=b.get("crot", 0),
                              text_color=b.get("text_color", "#000000"), bg=b.get("bg", False),
                              adj=b.get("adj"), flipH=b.get("flipH", False),
                              grad=b.get("grad"), n=b.get("n"),
                              pattern=b.get("pattern"),
                              stroke_w=b.get("stroke_w")))
        default_shape = spec.get("default_shape", "rect")
        for b in boxes:
            if b.get("shape") in (None, ""):
                b["shape"] = default_shape
        boxes.sort(key=lambda b: -(b["w"] * b["h"]))
        bg_elems = []

        for a in spec.get("arrows", []):
            if a.get("pts") and not a.get("from"):
                pts = [tuple(p) for p in a["pts"]]
            else:
                p1 = parse_anchor(a["from"], by_id)
                p2 = parse_anchor(a["to"], by_id)
                if a.get("pts"):
                    pts = [tuple(p) for p in a["pts"]]
                elif a.get("elbow", True):
                    pts = elbow_points(p1, p2, a.get("mode", ""))
                else:
                    pts = [p1, p2]
            add_polyline(slide, pts, to_emu, a.get("color", "#000000"), a.get("width", 1.0),
                         arrow=a.get("arrow", True), dash=a.get("dash", False))

        for pl in spec.get("polys", []):
            pts = [to_emu(px, py) for px, py in pl["pts"]]
            fb = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1.0)
            fb.add_line_segments(pts[1:], close=True)
            shp = fb.convert_to_shape()
            shp.shadow.inherit = False
            if pl.get("fill"):
                shp.fill.solid()
                shp.fill.fore_color.rgb = rgb(pl["fill"])
            else:
                shp.fill.background()
            if pl.get("stroke"):
                shp.line.color.rgb = rgb(pl["stroke"])
                shp.line.width = Pt(pl.get("width", 1.0))
                if pl.get("dash"):
                    shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            else:
                shp.line.fill.background()
            if pl.get("bg"):
                shp.name = "bg"
                bg_elems.append(shp._element)

        deferred_labels = []
        for b in boxes:
            shp_kind = b.get("shape", default_shape)
            x, y, w, h = b["x"], b["y"], b["w"], b["h"]
            rot = 0
            text_inside = True
            if shp_kind == "triangleR":
                rot = 90
                text_inside = False
                cx, cy = x + w / 2, y + h / 2
                x, y, w, h = cx - h / 2, cy - w / 2, h, w
            elif shp_kind == "triangleL":
                rot = 270
                text_inside = False
                cx, cy = x + w / 2, y + h / 2
                x, y, w, h = cx - h / 2, cy - w / 2, h, w
            elif b.get("crot"):
                rot = b["crot"]
                cx, cy = x, y
                x, y = cx - w / 2, cy - h / 2
            elif b.get("rot"):
                rot = b["rot"]
                cx, cy = x + w / 2, y + h / 2
                x, y, w, h = cx - h / 2, cy - w / 2, h, w
            shape_enum = SHAPE_MAP.get(shp_kind, MSO_SHAPE.ROUNDED_RECTANGLE)
            if shp_kind == "rect" and not spec.get("round", True):
                shape_enum = MSO_SHAPE.RECTANGLE
            shp = slide.shapes.add_shape(shape_enum, X(x), Y(y),
                                         Emu(int(w * scale * 914400)), Emu(int(h * scale * 914400)))
            shp.shadow.inherit = False
            if shp_kind == "rect" and b.get("adj") is None:
                try:
                    shp.adjustments[0] = 0.12
                except Exception:
                    pass
            if b.get("adj") is not None:
                try:
                    shp.adjustments[0] = b["adj"]
                except Exception:
                    pass
            if b.get("no_line"):
                shp.line.fill.background()
            else:
                stroke = b.get("stroke") or darker(b.get("fill", "#CCCCCC"))
                if str(stroke).upper() == "#FFFFFF":
                    stroke = darker(b.get("fill", "#AAAAAA"))
                shp.line.color.rgb = rgb(stroke)
                if b.get("stroke_w"):
                    shp.line.width = Pt(b["stroke_w"])
            if b.get("no_fill"):
                shp.fill.background()
            elif b.get("grad"):
                shp.fill.gradient()
                gs = shp.fill.gradient_stops
                cols = b["grad"]["colors"]
                for i in range(len(gs)):
                    s = gs[i]
                    s.color.rgb = rgb(cols[min(i, len(cols) - 1)])
                    s.position = i / (len(gs) - 1)
                try:
                    shp.fill.gradient_angle = b["grad"].get("angle", 0.0)
                except Exception:
                    pass
            elif b.get("pattern"):
                p = b["pattern"]
                shp.fill.patterned()
                shp.fill.pattern = MSO_PATTERN.SMALL_GRID
                shp.fill.fore_color.rgb = rgb(p.get("fg", "#FFFFFF"))
                shp.fill.back_color.rgb = rgb(p.get("bg", "#000000"))
            else:
                shp.fill.solid()
                shp.fill.fore_color.rgb = rgb(b.get("fill", "#FFFFFF"))
            if b.get("icon") == "stripes":
                n = b.get("n") or 4
                cw = b["w"] * 0.62
                ch = b["h"] * 0.80
                off = min(b["w"], b["h"]) * 0.14
                for k in range(n):
                    sx = b["x"] + (b["w"] - cw) / 2 - off * (n - 1) / 2 + off * k + cw * 0.28
                    sy = b["y"] + (b["h"] - ch) / 2 - off * (n - 1) / 2 + off * k
                    st = slide.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, X(sx), Y(sy),
                                                Emu(int(cw * scale * 914400)), Emu(int(ch * scale * 914400)))
                    st.shadow.inherit = False
                    st.line.color.rgb = rgb(b.get("stroke") or "#6C8EBF")
                    st.fill.solid()
                    st.fill.fore_color.rgb = rgb(b.get("fill") or "#FFFFFF")
                    st.adjustments[0] = 0.34
            if b.get("dash"):
                shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            if b.get("flipH"):
                xfrm = shp._element.spPr.find(qn("a:xfrm"))
                if xfrm is not None:
                    xfrm.set("flipH", "1")
            if rot:
                shp.rotation = rot
            if b.get("bg"):
                shp.name = "bg"
                bg_elems.append(shp._element)
            txt = b.get("text", "")
            if b.get("rot"):
                size = classify_size(b["h"], b["w"], txt, scale, b.get("size"))
            else:
                size = classify_size(b["w"], b["h"], txt, scale, b.get("size"))
            if txt and not text_inside:
                deferred_labels.append(dict(x=b["x"], y=b["y"], w=b["w"], h=b["h"],
                                            text=txt, size=b.get("size"),
                                            color=b.get("text_color", "#000000")))
            elif txt:
                set_text(shp, txt, size, color=b.get("text_color", "#000000"), font=font)

        img_path = img
        if not os.path.exists(img_path):
            img_path = os.path.join(os.path.dirname(sp_path), "..", img)
        for im_spec in spec.get("images", []):
            x, y, w, h = im_spec["rect"]
            crop = im_spec.get("crop", [x, y, w, h])
            tile = Image.open(img_path).convert("RGB").crop(
                (int(crop[0]), int(crop[1]), int(crop[0] + crop[2]), int(crop[1] + crop[3])))
            bio = io.BytesIO()
            tile.save(bio, "PNG")
            bio.seek(0)
            slide.shapes.add_picture(bio, X(x), Y(y),
                                     Emu(int(w * scale * 914400)), Emu(int(h * scale * 914400)))

        for lb in list(spec.get("labels", [])) + deferred_labels:
            tb = slide.shapes.add_textbox(X(lb["x"]), Y(lb["y"]),
                                          Emu(int(lb["w"] * scale * 914400)),
                                          Emu(int(lb["h"] * scale * 914400)))
            size = classify_size(lb["w"], lb["h"], lb["text"], scale, lb.get("size"))
            set_text(tb, lb["text"], size, align=lb.get("align", "c"),
                     color=lb.get("color", "#000000"), bold=lb.get("bold", False), font=font)
        if bg_elems:
            tree = slide.shapes._spTree
            for i, el in enumerate(bg_elems):
                tree.remove(el)
                tree.insert(2 + i, el)
        report.append(f"{sp_name}: {len(boxes)} shapes, "
                      f"{len(spec.get('arrows', []))} arrows  scale={scale:.4f}")
    prs.save(out_path)
    print("\n".join(report))
    print("saved:", out_path)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        specs = sys.argv[1:-1]
        out = sys.argv[-1]
    else:
        specs = sorted(glob.glob("spec_*.json"))
        out = "out.pptx"
    if not specs:
        raise SystemExit("no spec_*.json found (see reference/spec-format.md)")
    build(specs, out)
