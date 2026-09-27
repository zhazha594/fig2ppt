"""Detect filled shapes in diagram PNGs -> JSON + debug overlay.

fig2ppt v1.0.0 (2026-09-26) - packaged from the original _build pipeline.
Usage: python detect.py <image.png> <out.json> <debug_overlay.png>
           [--min-area 60] [--min-lum 140] [--keep-caption]
"""
import os
import sys
import json
import numpy as np
import cv2


def imread_u(path):
    return cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)


def imwrite_u(path, img):
    cv2.imencode(os.path.splitext(path)[1] or ".png", img)[1].tofile(path)


def lum(c):
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def hx(c):
    return "#%02X%02X%02X" % (int(c[0]), int(c[1]), int(c[2]))


def find_caption_top(rgb):
    """Cut below the last row with a wide (>=30px) colored run; ignores text links."""
    H, W = rgb.shape[:2]
    mx = rgb.max(axis=2).astype(np.int32)
    mn = rgb.min(axis=2).astype(np.int32)
    colored = (mx - mn > 60).astype(np.uint8)
    closed = cv2.morphologyEx(colored, cv2.MORPH_CLOSE, np.ones((1, 7), np.uint8))
    has_wide = np.zeros(H, bool)
    for r in range(H):
        row = closed[r]
        if row.sum() < 30:
            continue
        d = np.diff(np.concatenate(([0], row, [0])))
        starts = np.where(d == 1)[0]
        ends = np.where(d == -1)[0]
        if (ends - starts).max() >= 30:
            has_wide[r] = True
    rows = np.where(has_wide)[0]
    if len(rows) == 0:
        return H
    return min(H, int(rows[-1]) + 4)


def classify_shape(c_tl, c_tr, c_bl, c_br, te, be, le, re_):
    cf = (c_tl + c_tr + c_bl + c_br) / 4
    if cf < 0.18:
        if le > 0.72 and re_ < 0.4:
            return "triangleR"
        if re_ > 0.72 and le < 0.4:
            return "triangleL"
        if le > 0.72 and re_ > 0.72 and te < 0.6 and be < 0.6:
            return "triangleU"
        return "diamond"
    if cf < 0.5:
        return "ellipse_ish"
    return "rect"


def detect(path, min_area=60, min_lum=140, keep_caption=False):
    img = imread_u(path)
    if img is None:
        raise SystemExit(f"cannot read image: {path}")
    H, W = img.shape[:2]
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    if not keep_caption:
        cut = find_caption_top(rgb)
        if cut < H:
            rgb = rgb[:cut]
            img = img[:cut]
            H = cut
    q = (((rgb.astype(np.int32) + 4) // 8) * 8).clip(0, 255).astype(np.int32)
    codes = q[:, :, 0] * 65536 + q[:, :, 1] * 256 + q[:, :, 2]
    uniq, counts = np.unique(codes, return_counts=True)
    order = np.argsort(-counts)
    bg_code = int(uniq[order[0]])
    bg = ((bg_code >> 16) & 255, (bg_code >> 8) & 255, bg_code & 255)

    cands = []
    for idx in order[:40]:
        code = int(uniq[idx])
        cnt = int(counts[idx])
        if code == bg_code or cnt < 150:
            continue
        c = ((code >> 16) & 255, (code >> 8) & 255, code & 255)
        if lum(c) < min_lum:
            continue
        cands.append((code, c, cnt))
        if len(cands) >= 14:
            break

    boxes = []
    k_close = np.ones((9, 9), np.uint8)
    for code, c, cnt in cands:
        mask = (codes == code).astype(np.uint8) * 255
        closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k_close)
        dist = cv2.distanceTransform(closed, cv2.DIST_L2, 5)
        core = (dist >= 2.5).astype(np.uint8) * 255
        n, labels, stats, cent = cv2.connectedComponentsWithStats(core, 8)
        for i in range(1, n):
            x, y, w, h, area = stats[i]
            if area < min_area:
                continue
            x, y = max(0, x - 3), max(0, y - 3)
            w, h = min(W - x, w + 6), min(H - y, h + 6)
            if w < 12 or h < 10:
                continue
            win = closed[y:y + h, x:x + w] > 0
            ratio = float(win.sum()) / (w * h)
            if ratio < 0.42:
                continue
            # shape guess
            px = max(2, int(min(w, h) * 0.20))
            c_tl = float(win[:px, :px].mean())
            c_tr = float(win[:px, -px:].mean())
            c_bl = float(win[-px:, :px].mean())
            c_br = float(win[-px:, -px:].mean())
            te = float(win[:px, :].mean())
            be = float(win[-px:, :].mean())
            le = float(win[:, :px].mean())
            re_ = float(win[:, -px:].mean())
            shape = classify_shape(c_tl, c_tr, c_bl, c_br, te, be, le, re_)
            stats_row = dict(x=int(x), y=int(y), w=int(w), h=int(h), area=int(area),
                             fill=hx(c), stroke=None, shape=shape, ratio=round(ratio, 2),
                             n=int(win.sum()),
                             te=round(te, 2), be=round(be, 2), le=round(le, 2), re=round(re_, 2),
                             c_tl=round(c_tl, 2), c_tr=round(c_tr, 2), c_bl=round(c_bl, 2), c_br=round(c_br, 2))
            # stroke color: ring around bbox
            x0, y0 = max(0, x - 3), max(0, y - 3)
            x1, y1 = min(W, x + w + 3), min(H, y + h + 3)
            ring = rgb[y0:y1, x0:x1].reshape(-1, 3)
            rq = q[y0:y1, x0:x1].reshape(-1, 3)
            rcodes = rq[:, 0] * 65536 + rq[:, 1] * 256 + rq[:, 2]
            keep = (rcodes != code) & (rcodes != bg_code)
            stroke = None
            if keep.any():
                u2, c2 = np.unique(rcodes[keep], return_counts=True)
                j = np.argsort(-c2)[0]
                cc = int(u2[j])
                stroke = ((cc >> 16) & 255, (cc >> 8) & 255, cc & 255)
            stats_row['stroke'] = hx(stroke) if stroke else None
            boxes.append(stats_row)

    # pass 2: enclosed near-white regions (white-filled boxes with colored borders)
    near_white = ((rgb.min(axis=2) > 235)).astype(np.uint8)
    n2, lab2, st2, _ = cv2.connectedComponentsWithStats(near_white, 4)
    border_ids = set(lab2[0].tolist()) | set(lab2[-1].tolist()) | set(lab2[:, 0].tolist()) | set(lab2[:, -1].tolist())
    for i in range(1, n2):
        x, y, w, h, area = st2[i]
        if i in border_ids or area < 200 or w < 14 or h < 12:
            continue
        if w * h > 300000:
            continue
        ratio = area / (w * h)
        if ratio < 0.85:
            continue
        x0, y0 = max(0, x - 3), max(0, y - 3)
        x1, y1 = min(W, x + w + 3), min(H, y + h + 3)
        ring = rgb[y0:y1, x0:x1].reshape(-1, 3)
        sat = ring.max(axis=1).astype(np.int32) - ring.min(axis=1).astype(np.int32)
        stroke = None
        if (sat > 25).any():
            sel = ring[sat > 25]
            med = np.median(sel, axis=0)
            stroke = tuple(int(v) for v in med)
        px = max(2, int(min(w, h) * 0.20))
        win = near_white[y:y + h, x:x + w].astype(bool)
        c_tl = float(win[:px, :px].mean()); c_tr = float(win[:px, -px:].mean())
        c_bl = float(win[-px:, :px].mean()); c_br = float(win[-px:, -px:].mean())
        boxes.append(dict(x=int(x), y=int(y), w=int(w), h=int(h), area=int(area),
                          fill="#FFFFFF", stroke=hx(stroke) if stroke else None,
                          shape="rect", ratio=round(ratio, 2), n=int(area),
                          te=1.0, be=1.0, le=1.0, re=1.0,
                          c_tl=round(c_tl, 2), c_tr=round(c_tr, 2), c_bl=round(c_bl, 2), c_br=round(c_br, 2),
                          white=True))

    # sort into reading order (row bands of 18px)
    boxes.sort(key=lambda b: (round(b['y'] / 18), b['x']))
    # drop near-duplicates: same-ish bbox, keep bigger
    out = []
    for b in boxes:
        dup = False
        for o in out:
            if o['fill'] != b['fill']:
                continue
            ix = max(0, min(b['x'] + b['w'], o['x'] + o['w']) - max(b['x'], o['x']))
            iy = max(0, min(b['y'] + b['h'], o['y'] + o['h']) - max(b['y'], o['y']))
            inter = ix * iy
            if inter > 0.8 * min(b['w'] * b['h'], o['w'] * o['h']):
                dup = True
                break
        if not dup:
            out.append(b)
    for i, b in enumerate(out):
        b['id'] = i

    # debug overlay
    ov = img.copy()
    for b in out:
        cv2.rectangle(ov, (b['x'], b['y']), (b['x'] + b['w'], b['y'] + b['h']), (0, 0, 255), 2)
        cv2.putText(ov, str(b['id']), (b['x'] + 3, b['y'] + 18), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (255, 0, 255), 2)
    return dict(image=path.split('/')[-1].split('\\')[-1], W=W, H=H, bg=hx(bg), boxes=out), ov


def parse_args(argv):
    pos, kw = [], {"min_area": 60, "min_lum": 140, "keep_caption": False}
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--min-area":
            kw["min_area"] = float(argv[i + 1]); i += 2
        elif a == "--min-lum":
            kw["min_lum"] = float(argv[i + 1]); i += 2
        elif a == "--keep-caption":
            kw["keep_caption"] = True; i += 1
        else:
            pos.append(a); i += 1
    return pos, kw


if __name__ == "__main__":
    pos, kw = parse_args(sys.argv[1:])
    if len(pos) < 3:
        raise SystemExit(__doc__)
    src, oj, ov = pos[0], pos[1], pos[2]
    data, overlay = detect(src, **kw)
    with open(oj, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    imwrite_u(ov, overlay)
    print(f"{data['image']}: {len(data['boxes'])} boxes  {data['W']}x{data['H']}")
    for b in data['boxes']:
        print(f"  {b['id']:>3}  x={b['x']:>4} y={b['y']:>4} w={b['w']:>3} h={b['h']:>3} "
              f"{b['fill']} {b['shape']:<11} r={b['ratio']:<4} le={b['le']} re={b['re']}")
