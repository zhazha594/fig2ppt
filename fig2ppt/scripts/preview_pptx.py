"""Render a PPTX to PNG per slide with matplotlib (no PowerPoint needed).

fig2ppt v1.0.0 (2026-09-26) - packaged from the original _build pipeline.
Usage: python preview_pptx.py <file.pptx> [out_prefix]
       out_prefix with %d formats per slide (default prev%d.png);
       multi-slide pptx without %d gets _<n> inserted automatically.
"""
import io
import os
import sys
import math
import numpy as np
from PIL import Image as PILImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL
from pptx.oxml.ns import qn
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Ellipse, Rectangle
import matplotlib.transforms as mtransforms

EMU_PER_IN = 914400.0
DPI = 130


def solid_rgb(fill):
    try:
        if fill.type == MSO_FILL.SOLID:
            c = fill.fore_color.rgb
            return "#%02X%02X%02X" % (c[0], c[1], c[2])
    except Exception:
        pass
    return None


def grad_colors(shape):
    """Return list of (pos, '#RRGGBB') stops if the shape has a gradient fill."""
    try:
        spPr = shape._element.spPr
    except Exception:
        return None
    gf = spPr.find(qn("a:gradFill"))
    if gf is None:
        return None
    stops = []
    gsLst = gf.find(qn("a:gsLst"))
    if gsLst is None:
        return None
    for gs in gsLst.findall(qn("a:gs")):
        c = gs.find(qn("a:srgbClr"))
        if c is not None:
            v = c.get("val")
            stops.append((int(gs.get("pos", "0")) / 100000.0, "#" + v))
    return stops or None



def pattern_colors(shape):
    """Return (fg, bg) hex strings if the shape has a pattern fill."""
    try:
        spPr = shape._element.spPr
    except Exception:
        return None
    pf = spPr.find(qn("a:pattFill"))
    if pf is None:
        return None

    def clr(el):
        c = el.find(qn("a:srgbClr")) if el is not None else None
        return "#" + c.get("val") if c is not None else None

    fg = clr(pf.find(qn("a:fgClr")))
    bg = clr(pf.find(qn("a:bgClr")))
    if fg is None or bg is None:
        return None
    return (fg, bg)


def line_style(shape):
    try:
        ln = shape.line
        color = solid_rgb(ln.fill)
        if color is None:
            return None, 0, False
        w = ln.width.pt if ln.width is not None else 1.0
        dash = False
        try:
            dash = ln.dash_style is not None
        except Exception:
            dash = False
        return color, w, dash
    except Exception:
        return None, 0, False


def poly_for(kind, x0, y0, x1, y1, adj=None):
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    w, h = x1 - x0, y1 - y0
    if kind == "ROUNDED_RECTANGLE":
        r = min(w, h) * (adj if adj is not None else 0.16667)
        r = max(0.0, min(r, w / 2, h / 2))
        if r > 0.01:
            pts = []
            for ccx, ccy, a0 in ((x0 + r, y0 + r, 180), (x1 - r, y0 + r, 270),
                                 (x1 - r, y1 - r, 0), (x0 + r, y1 - r, 90)):
                for k in range(0, 9):
                    a = math.radians(a0 + k * 90 / 8)
                    pts.append((ccx + r * math.cos(a), ccy + r * math.sin(a)))
            return pts
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    if kind == "RECTANGLE":
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    if kind == "DIAMOND":
        return [(cx, y0), (x1, cy), (cx, y1), (x0, cy)]
    if kind == "ISOSCELES_TRIANGLE":
        return [(cx, y0), (x1, y1), (x0, y1)]
    if kind == "HEXAGON":
        i = min(w, h) * (adj if adj is not None else 0.25)
        return [(x0 + i, y0), (x1 - i, y0), (x1, cy), (x1 - i, y1), (x0 + i, y1), (x0, cy)]
    if kind == "PARALLELOGRAM":
        s = min(w, h) * (adj if adj is not None else 0.28)
        return [(x0 + s, y0), (x1, y0), (x1 - s, y1), (x0, y1)]
    if kind == "TRAPEZOID":
        i = min(w, h) * (adj if adj is not None else 0.25)
        return [(x0 + i, y0), (x1 - i, y0), (x1, y1), (x0, y1)]
    if kind == "PENTAGON":
        i = min(w, h) * (adj if adj is not None else 0.25)
        return [(x0 + i, y0), (x1 - i, y0), (x1, cy), (cx, y1), (x0, cy)]
    if kind == "CAN":
        return [(x0, y0 + h * 0.12), (x1, y0 + h * 0.12), (x1, y1 - h * 0.12), (x0, y1 - h * 0.12)]
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]


def freeform_points(shape):
    sp = shape._element
    paths = []
    for path in sp.iter(qn("a:path")):
        pw = float(path.get("w", shape.width))
        ph = float(path.get("h", shape.height))
        pts = []
        for child in path:
            tag = child.tag
            if tag in (qn("a:moveTo"), qn("a:lnTo")):
                pt = child.find(qn("a:pt"))
                if pt is None:
                    continue
                px = float(pt.get("x"))
                py = float(pt.get("y"))
                pts.append((shape.left + px * (shape.width / pw if pw else 1),
                            shape.top + py * (shape.height / ph if ph else 1)))
        if pts:
            paths.append(pts)
    return paths


def has_arrow(shape):
    ln = shape._element.find(qn("p:spPr") + "/" + qn("a:ln"))
    if ln is None:
        ln = shape._element.find(".//" + qn("a:ln"))
    if ln is None:
        return False
    e = ln.find(qn("a:tailEnd"))
    return e is not None and e.get("type") not in (None, "none")


def draw_text(ax, shape, x0, y0, x1, y1, rot=0):
    tf = shape.text_frame
    if not tf.text.strip():
        return
    size = None
    color = "#000000"
    bold = False
    align = "c"
    fam = None
    for p in tf.paragraphs:
        for r in p.runs:
            if size is None and r.font.size is not None:
                size = r.font.size.pt
            if r.font.name:
                fam = r.font.name
            try:
                if r.font.color is not None and r.font.color.rgb is not None:
                    c = r.font.color.rgb
                    color = "#%02X%02X%02X" % (c[0], c[1], c[2])
            except Exception:
                pass
            bold = bool(r.font.bold)
            if p.alignment is not None:
                align = str(p.alignment).split(" ")[0].split("(")[0].lower()
        if size is not None:
            break
    if size is None:
        size = 8
    lines = tf.text.split("\n")
    n = len(lines)
    lh = size * 1.22 / 72.0
    cy = (y0 + y1) / 2
    ty = cy + (n - 1) * lh / 2
    if "center" in align or align == "c":
        tx, ha = (x0 + x1) / 2, "center"
    elif "left" in align or align == "l":
        tx, ha = x0 + 0.02, "left"
    else:
        tx, ha = x1 - 0.02, "right"
    ax.text(tx, ty, tf.text, fontsize=size, ha=ha, va="center", color=color,
            family=fam or "Times New Roman", weight="bold" if bold else "normal",
            linespacing=1.22,
            rotation=(-rot) % 360, rotation_mode="anchor", zorder=50)


def render(pptx_path, out_prefix):
    prs = Presentation(pptx_path)
    W = prs.slide_width / EMU_PER_IN
    H = prs.slide_height / EMU_PER_IN
    slides = list(prs.slides)
    n_slides = len(slides)
    outs = []
    for idx, slide in enumerate(slides, 1):
        fig = plt.figure(figsize=(W, H), dpi=DPI)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, W)
        ax.set_ylim(H, 0)
        ax.axis("off")
        ax.add_patch(Rectangle((0, 0), W, H, facecolor="white", edgecolor="none", zorder=0))
        for shp in slide.shapes:
            x0 = shp.left / EMU_PER_IN
            y0 = shp.top / EMU_PER_IN
            x1 = x0 + shp.width / EMU_PER_IN
            y1 = y0 + shp.height / EMU_PER_IN
            st = shp.shape_type
            if st == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = shp.image.blob
                    arr = np.asarray(PILImage.open(io.BytesIO(blob)).convert("RGB"))
                    ax.imshow(arr, extent=[x0, x1, y1, y0], zorder=15, aspect="auto")
                except Exception:
                    pass
                continue
            if st == MSO_SHAPE_TYPE.FREEFORM:
                fc = solid_rgb(shp.fill)
                groups = freeform_points(shp)
                if fc and groups:
                    pts = [(p[0] / EMU_PER_IN, p[1] / EMU_PER_IN) for p in groups[0]]
                    ax.add_patch(Polygon(pts, closed=True, facecolor=fc, edgecolor="none",
                                         zorder=1 if shp.name == "bg" else 10))
                    continue
                lc, lw, dash = line_style(shp)
                for pts in groups:
                    xs = [p[0] / EMU_PER_IN for p in pts]
                    ys = [p[1] / EMU_PER_IN for p in pts]
                    ax.plot(xs, ys, color=lc or "#000000", lw=max(lw, 0.6),
                            ls=(0, (3, 2)) if dash else "-", zorder=20,
                            solid_capstyle="round")
                    if has_arrow(shp) and len(xs) >= 2:
                        (px, py), (qx, qy) = (xs[-2], ys[-2]), (xs[-1], ys[-1])
                        ang = math.atan2(qy - py, qx - px)
                        head = max(0.055, lw / 72.0 * 5)
                        for s in (-1, 1):
                            a2 = ang + math.pi + s * 0.42
                            ax.plot([qx, qx + head * math.cos(a2)], [qy, qy + head * math.sin(a2)],
                                    color=lc or "#000000", lw=max(lw, 0.6), zorder=20)
                continue
            if st == MSO_SHAPE_TYPE.TEXT_BOX:
                draw_text(ax, shp, x0, y0, x1, y1)
                continue
            kind = None
            try:
                kind = str(shp.auto_shape_type).split(" ")[0]
            except Exception:
                kind = "RECTANGLE"
            fc = solid_rgb(shp.fill)
            grad = grad_colors(shp) if fc is None else None
            patt = pattern_colors(shp) if fc is None else None
            lc, lw, dash = line_style(shp)
            rot = shp.rotation or 0
            z = 1 if shp.name == "bg" else 10
            patch = None
            if kind == "OVAL":
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                patch = Ellipse((cx, cy), x1 - x0, y1 - y0, facecolor=fc or "none",
                                edgecolor=lc or "none", lw=max(lw, 0.6) if lc else 0, zorder=z)
                ax.add_patch(patch)
            else:
                adj = None
                try:
                    adj = shp.adjustments[0]
                except Exception:
                    adj = None
                pts = poly_for(kind, x0, y0, x1, y1, adj)
                try:
                    xfrm = shp._element.spPr.find(qn("a:xfrm"))
                    if xfrm is not None and xfrm.get("flipH") == "1":
                        pts = [(x0 + x1 - px, py) for (px, py) in pts]
                except Exception:
                    pass
                patch = Polygon(pts, closed=True, facecolor=fc or "none",
                                edgecolor=lc or "none", lw=max(lw, 0.6) if lc else 0,
                                zorder=z, joinstyle="miter")
                if rot:
                    # data space has y increasing downward, so a positive OOXML
                    # rotation (clockwise on screen) is a positive algebraic angle here
                    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                    patch.set_transform(mtransforms.Affine2D().rotate_deg_around(cx, cy, rot) + ax.transData)
                ax.add_patch(patch)
            if grad and patch is not None:
                cols = [matplotlib.colors.to_rgb(c) for (_, c) in grad]
                pos = [p for (p, _) in grad]
                n = 256
                arr = np.zeros((1, n, 3))
                for i in range(n):
                    t = i / (n - 1)
                    for k in range(len(cols) - 1):
                        if pos[k] <= t <= pos[k + 1] or k == len(cols) - 2:
                            span = max(1e-6, pos[k + 1] - pos[k])
                            u = min(1.0, max(0.0, (t - pos[k]) / span))
                            arr[0, i] = [cols[k][j] + (cols[k + 1][j] - cols[k][j]) * u for j in range(3)]
                            break
                im = ax.imshow(arr, extent=[x0, x1, y1, y0], zorder=z + 0.2, aspect="auto")
                im.set_clip_path(patch)
            if patt and patch is not None:
                fg_rgb = matplotlib.colors.to_rgb(patt[0])
                bg_rgb = matplotlib.colors.to_rgb(patt[1])
                tile = np.zeros((6, 6, 3))
                tile[:, :] = bg_rgb
                tile[0, :] = fg_rgb
                tile[:, 0] = fg_rgb
                im = ax.imshow(tile, extent=[x0, x1, y1, y0], zorder=z + 0.2, aspect="auto",
                               interpolation="nearest")
                im.set_clip_path(patch)
            draw_text(ax, shp, x0, y0, x1, y1, rot=rot)
        if "%" in out_prefix:
            out = out_prefix % idx
        elif n_slides > 1:
            base, ext = os.path.splitext(out_prefix)
            out = f"{base}_{idx}{ext or '.png'}"
        else:
            out = out_prefix
        fig.savefig(out, dpi=DPI, facecolor="white")
        plt.close(fig)
        outs.append(out)
    return outs


if __name__ == "__main__":
    pptx = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else "prev%d.png"
    for o in render(pptx, prefix):
        print("wrote", o)
