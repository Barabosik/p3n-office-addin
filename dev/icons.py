# Builds the add-in's icons from dev/logo-512.png (the same source and tile as the browser extension's dev/icons.py).
#   python dev/icons.py [preview]
# The logo's pen body is transparent with a black outline, which vanishes on Word's dark ribbon (the Black and
# Dark Gray Office themes). So every icon Office draws sits on a light rounded tile that reads on light and dark
# ribbons. Everything Office draws (the ribbon button at 16, 32 and 80 px, and the add-in list icon at 32 and
# 64 px, the same icon at two DPIs) shows only the pen: the swooshes below it turn to noise that small. The
# 128 px icon and the 300 px AppSource logo show the whole logo. logo.png is the plain logo for the pane's
# header on light paper; logo-dark.png, from dev/logo-512-dark.png (p3n.si's dark logo: the black line art in the
# dark theme's ink), is the same on the dark pane.
import sys
from pathlib import Path
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
PAPER, RULE = (251, 250, 247, 255), (230, 225, 216, 255)


def tile(logo, size, pad):
    """The logo on a rounded light tile with a hairline edge, drawn at 8x and scaled down."""
    big = size * 8
    out = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    ImageDraw.Draw(out).rounded_rectangle((0, 0, big - 1, big - 1), radius=int(big * 0.22), fill=PAPER, outline=RULE, width=max(8, big // 64))
    inner = int(big * (1 - 2 * pad))
    out.alpha_composite(logo.resize((inner, inner), Image.LANCZOS), ((big - inner) // 2, (big - inner) // 2))
    return out.resize((size, size), Image.LANCZOS)


def square(img):
    img = img.crop(img.getbbox())
    side = max(img.size)
    sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    sq.alpha_composite(img, ((side - img.width) // 2, (side - img.height) // 2))
    return sq


src = Image.open(ROOT / "dev" / "logo-512.png").convert("RGBA")
logo = square(src)
pen = square(src.crop((105, 30, 470, 395)))  # the pen and its nib, without the swooshes

out_dir = ROOT / ("dev/out/icons" if "preview" in sys.argv else "src/assets")
out_dir.mkdir(parents=True, exist_ok=True)
for size, art, pad in ((16, pen, 0.08), (32, pen, 0.1), (64, pen, 0.1), (80, pen, 0.12), (128, logo, 0.12)):
    tile(art, size, pad).save(out_dir / f"icon-{size}.png", optimize=True)
logo.resize((128, 128), Image.LANCZOS).save(out_dir / "logo.png", optimize=True)
dark = Image.open(ROOT / "dev" / "logo-512-dark.png").convert("RGBA")
square(dark).resize((128, 128), Image.LANCZOS).save(out_dir / "logo-dark.png", optimize=True)
# AppSource wants a logo between 216 and 350 px square.
store_dir = ROOT / ("dev/out/icons" if "preview" in sys.argv else "store")
tile(logo, 300, 0.12).save(store_dir / "logo-300.png", optimize=True)
print("icons written to", out_dir.relative_to(ROOT))
