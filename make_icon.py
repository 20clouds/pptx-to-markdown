"""Draw the droplet's icon and write it as an .icns file: make_icon.py OUT.icns"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

S = 1024


def font(size):
    try:
        return ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", size, index=1)
    except Exception:
        return ImageFont.load_default(size=size)


def draw() -> Image.Image:
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    # vertical PowerPoint-orange gradient inside a macOS-style rounded square
    grad = Image.new("RGBA", (S, S))
    gd = ImageDraw.Draw(grad)
    top, bottom = (232, 106, 64), (183, 56, 30)
    for y in range(S):
        t = y / (S - 1)
        gd.line([(0, y), (S, y)], fill=tuple(int(a + (b - a) * t) for a, b in zip(top, bottom)) + (255,))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle([100, 100, S - 100, S - 100], radius=185, fill=255)
    im.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(im)
    # a slide card
    d.rounded_rectangle([210, 230, 814, 590], radius=36, fill=(255, 255, 255, 255))
    d.rectangle([260, 280, 560, 312], fill=(183, 56, 30, 255))
    for i, w in enumerate((440, 380, 410)):
        d.rectangle([260, 350 + i * 60, 260 + w, 372 + i * 60], fill=(200, 200, 205, 255))
    d.rounded_rectangle([620, 340, 764, 540], radius=14, fill=(255, 196, 160, 255))
    # arrow (drawn, since Helvetica Bold has no arrow glyph) + MD
    f = font(210)
    text_w = d.textlength("MD", font=f)
    arrow_w, gap = 170, 50
    x0 = (S - (arrow_w + gap + text_w)) / 2
    cy = 640 + 115
    white = (255, 255, 255, 255)
    d.rectangle([x0, cy - 22, x0 + arrow_w - 70, cy + 22], fill=white)
    d.polygon([(x0 + arrow_w - 90, cy - 70), (x0 + arrow_w, cy), (x0 + arrow_w - 90, cy + 70)], fill=white)
    d.text((x0 + arrow_w + gap, 640), "MD", font=f, fill=white)
    return im


def main(out: Path):
    base = draw()
    tmp = Path(tempfile.mkdtemp())
    iconset = tmp / "icon.iconset"
    iconset.mkdir()
    for size in (16, 32, 128, 256, 512):
        base.resize((size, size), Image.LANCZOS).save(iconset / f"icon_{size}x{size}.png")
        base.resize((size * 2, size * 2), Image.LANCZOS).save(iconset / f"icon_{size}x{size}@2x.png")
    subprocess.run(["/usr/bin/iconutil", "-c", "icns", str(iconset), "-o", str(out)], check=True)
    shutil.rmtree(tmp)


if __name__ == "__main__":
    main(Path(sys.argv[1]))
