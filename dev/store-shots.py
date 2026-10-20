# Builds the AppSource screenshots, 1366x768 and under 1024 KB each, from window captures of real Word and
# Excel, and the two the GitHub profile uses:   python dev/store-shots.py
# They come from `SHOTS=1 node dev/live.mjs word|excel` (dev/out/raw-*.png): a 1366x768 window at 100%
# scaling, or that times the display's scaling. The account picture in the title bar is painted over with
# the bar's own colour. The README uses the same images.
import sys
from pathlib import Path
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "dev" / "out"
W, H = 1366, 768


def paint_avatar(im):
    """The round account picture between the search box and the minimize button, in the title bar."""
    k = im.width / W
    # The title bar is 48 px tall at 100%; below it, the Comments and Share buttons must stay out of the box.
    box = (int(im.width - 250 * k), int(4 * k), int(im.width - 135 * k), int(44 * k))
    region = im.crop(box)
    bar = max(region.getcolors(region.width * region.height), key=lambda c: c[0])[1]
    hits = [(x, y) for x in range(region.width) for y in range(region.height)
            if sum(abs(a - b) for a, b in zip(region.getpixel((x, y)), bar)) > 60]
    if not hits:
        return im
    xs, ys = [x for x, _ in hits], [y for _, y in hits]
    pad = int(3 * k)
    ImageDraw.Draw(im).rectangle((box[0] + min(xs) - pad, box[1] + min(ys) - pad, box[0] + max(xs) + pad, box[1] + max(ys) + pad), fill=bar)
    return im


def to_store(im, name):
    im = im.convert("RGB")
    if im.size != (W, H):
        assert abs(im.width / im.height - W / H) < 0.01, f"{name}: {im.size} isn't 1366:768"
        im = im.resize((W, H), Image.LANCZOS)
    path = ROOT / "store" / name
    im.save(path, optimize=True)
    size = path.stat().st_size
    assert size <= 1024 * 1024, f"{name} is {size} bytes, over 1024 KB"
    print(f"store/{name}: {im.size}, {size // 1024} KB")


# Word, Check.
raw = OUT / "raw-word-check.png"
if not raw.exists():
    sys.exit(f"{raw.relative_to(ROOT)} is missing: run SHOTS=1 node dev/live.mjs word first")
to_store(paint_avatar(Image.open(raw).convert("RGB")), "screenshot-word-check.png")

# After Replace for AppSource; before it (the score, the rewrite and the Replace button) as
# screenshot-word.png and screenshot-excel.png, which the GitHub profile's composite image uses.
for app in ("word", "excel"):
    for state, name in (("replaced", f"screenshot-{app}-replace.png"), ("rewrite", f"screenshot-{app}.png")):
        raw = OUT / f"raw-{app}-{state}.png"
        if not raw.exists():
            sys.exit(f"{raw.relative_to(ROOT)} is missing: run SHOTS=1 node dev/live.mjs {app} first")
        to_store(paint_avatar(Image.open(raw).convert("RGB")), name)
