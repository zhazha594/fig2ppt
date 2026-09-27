"""fig2ppt smoke test: run detect -> build -> preview on the bundled sample.

Usage: D:\\Python\\python.exe run_selftest.py
Exits 0 and prints SELFTEST OK if the full chain produces sane outputs.
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "..", "scripts")
PY = sys.executable
IMG = os.path.join(HERE, "decb7beaf6cc6b072426b0d8e300a180.png")
GEO = os.path.join(HERE, "decb7beaf6cc6b072426b0d8e300a180.json")
DBG = os.path.join(HERE, "dbg_selftest.png")
SPEC = os.path.join(HERE, "spec_fig7.json")
PPTX = os.path.join(HERE, "selftest_out.pptx")
PREV = os.path.join(HERE, "selftest_prev%d.png")


def run(cmd):
    print("+", " ".join(os.path.basename(c) if c.endswith(".py") else c for c in cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=HERE)
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise SystemExit("SELFTEST FAILED (see output above)")
    return r.stdout


def main():
    out = run([PY, os.path.join(SCRIPTS, "detect.py"), IMG, GEO, DBG])
    geo = json.load(open(GEO, encoding="utf-8"))
    assert len(geo["boxes"]) >= 10, f"too few boxes: {len(geo['boxes'])}"
    print(f"  detect: {len(geo['boxes'])} boxes")

    out = run([PY, os.path.join(SCRIPTS, "build_pptx.py"), SPEC, PPTX])
    print(" ", out.strip().replace(chr(10), chr(10) + "  "))
    assert os.path.getsize(PPTX) > 10000, "pptx suspiciously small"
    assert "WARN" not in out, "unexpected WARN from known-good spec"

    from pptx import Presentation
    n_shapes = len(Presentation(PPTX).slides[0].shapes)
    assert n_shapes >= 20, f"too few shapes on slide: {n_shapes}"
    print(f"  slide shapes: {n_shapes}")

    run([PY, os.path.join(SCRIPTS, "preview_pptx.py"), PPTX, PREV])
    prev = PREV % 1
    assert os.path.getsize(prev) > 20000, "preview suspiciously small"
    print(f"  preview: {os.path.basename(prev)} {os.path.getsize(prev)//1024}KB")
    print("SELFTEST OK")


if __name__ == "__main__":
    main()
