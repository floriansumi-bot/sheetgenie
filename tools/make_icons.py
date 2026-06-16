"""Generate SheetGenie PNG app icons (maskable-safe, full-bleed) from code.

Produces public/icons/icon-192.png, icon-512.png, apple-touch-icon-180.png to
complement the scalable icon.svg, for broad Android launcher + iOS home-screen
support. Run:  .venv\\Scripts\\python.exe tools\\make_icons.py
"""

import os
from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "public", "icons")
VIOLET = (139, 92, 246, 255)
VIOLET_LT = (167, 139, 250, 255)


def _sparkle(d, cx, cy, r):
    # Soft glow, then a 4-point violet star with a white core.
    d.ellipse([cx - r * 1.15, cy - r * 1.15, cx + r * 1.15, cy + r * 1.15], fill=(124, 92, 246))
    d.polygon([(cx, cy - r), (cx + r * 0.30, cy), (cx, cy + r), (cx - r * 0.30, cy)], fill=(237, 233, 254))
    d.polygon([(cx - r, cy), (cx, cy - r * 0.30), (cx + r, cy), (cx, cy + r * 0.30)], fill=(237, 233, 254))
    d.ellipse([cx - r * 0.13, cy - r * 0.13, cx + r * 0.13, cy + r * 0.13], fill=(255, 255, 255))


def _render(size):
    # Drawn with OPAQUE colors (ImageDraw overwrites rather than alpha-composites),
    # then downscaled for anti-aliasing.
    S = 1024
    img = Image.new("RGB", (S, S), (12, 11, 19))
    d = ImageDraw.Draw(img)

    # Full-bleed diagonal dark-violet -> near-black background (maskable-safe).
    top, bot = (36, 28, 64), (12, 11, 19)
    for y in range(S):
        t = y / S
        d.line([(0, y), (S, y)], fill=(
            int(top[0] * (1 - t) + bot[0] * t),
            int(top[1] * (1 - t) + bot[1] * t),
            int(top[2] * (1 - t) + bot[2] * t),
        ))

    # White spreadsheet "sheet" (content within the maskable safe zone).
    cw, ch = int(S * 0.50), int(S * 0.58)
    cx, cy = (S - cw) // 2, (S - ch) // 2 - int(S * 0.02)
    rad = int(S * 0.055)
    d.rounded_rectangle([cx, cy, cx + cw, cy + ch], radius=rad, fill=(247, 247, 251))

    # Violet header band (rounded top, flat bottom).
    hb = int(ch * 0.17)
    d.rounded_rectangle([cx, cy, cx + cw, cy + hb], radius=rad, fill=VIOLET[:3])
    d.rectangle([cx, cy + hb - rad, cx + cw, cy + hb], fill=VIOLET[:3])

    # Grey gridlines across the body so it reads as a spreadsheet.
    body_top, body_bot = cy + hb, cy + ch
    grid = (206, 206, 218)
    gw = max(2, S // 300)
    for i in (1, 2):
        vx = cx + int(cw * i / 3)
        d.line([(vx, body_top), (vx, body_bot)], fill=grid, width=gw)
    for j in (1, 2, 3):
        hy = body_top + int((body_bot - body_top) * j / 4)
        d.line([(cx, hy), (cx + cw, hy)], fill=grid, width=gw)

    _sparkle(d, cx + cw - int(S * 0.01), cy + ch - int(S * 0.01), int(S * 0.15))

    return img.resize((size, size), Image.LANCZOS)


def main():
    os.makedirs(OUT, exist_ok=True)
    targets = {"icon-192.png": 192, "icon-512.png": 512, "apple-touch-icon-180.png": 180}
    for name, size in targets.items():
        # RGB (no alpha) — full-bleed, good for both maskable and iOS touch icons.
        _render(size).save(os.path.join(OUT, name))
        print("wrote", name, size)


if __name__ == "__main__":
    main()
