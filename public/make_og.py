#!/usr/bin/env python
"""Generate og-image.png (1200x630) for SheetGenie link previews.

On-brand: dark slate background with a violet->indigo accent, the SheetGenie
name, the spreadsheet+sparkle mark, and the tagline. Rendered at 2x and
downsampled for crisp anti-aliasing. Pure-Pillow (no SVG rasterizer needed).

Run: .venv\\Scripts\\python.exe public\\make_og.py
"""
import math
from PIL import Image, ImageDraw, ImageFont, ImageFilter

# ---- brand constants -------------------------------------------------------
W, H = 1200, 630
S = 2  # supersample factor
BG_TOP = (30, 27, 46)      # #1e1b2e
BG_BOT = (15, 14, 23)      # #0f0e17
ACCENT_1 = (139, 92, 246)  # #8b5cf6
ACCENT_2 = (99, 102, 241)  # #6366f1
ACCENT_LT = (167, 139, 250)  # #a78bfa
TEXT = (236, 233, 245)     # #ece9f5
TEXT_DIM = (163, 159, 181)  # #a39fb5
GRID_LINE = (139, 92, 246)  # grid tint
CARD_STROKE = (59, 58, 82)  # #3b3a52
FONTS = "C:/Windows/Fonts/"


def lerp(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def vgrad(size, top, bottom):
    """Vertical gradient image."""
    w, h = size
    grad = Image.new("RGB", (1, h))
    px = grad.load()
    for y in range(h):
        px[0, y] = lerp(top, bottom, y / max(1, h - 1))
    return grad.resize((w, h))


def diag_grad_rgba(size, c1, c2):
    """Diagonal (TL->BR) gradient, returned as RGBA."""
    w, h = size
    base = Image.new("RGB", (w, h))
    px = base.load()
    denom = max(1, (w - 1) + (h - 1))
    for y in range(h):
        for x in range(w):
            px[x, y] = lerp(c1, c2, (x + y) / denom)
    return base.convert("RGBA")


def font(name, px):
    return ImageFont.truetype(FONTS + name, px)


# ---- canvas ----------------------------------------------------------------
w, h = W * S, H * S
img = vgrad((w, h), BG_TOP, BG_BOT).convert("RGBA")
draw = ImageDraw.Draw(img, "RGBA")

# subtle spreadsheet grid (matches the app's 28px grid backdrop)
grid_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
gd = ImageDraw.Draw(grid_layer)
cell = 28 * S
ga = 14  # alpha
for x in range(0, w, cell):
    gd.line([(x, 0), (x, h)], fill=GRID_LINE + (ga,), width=S)
for y in range(0, h, cell):
    gd.line([(0, y), (w, y)], fill=GRID_LINE + (ga,), width=S)
img = Image.alpha_composite(img, grid_layer)
draw = ImageDraw.Draw(img, "RGBA")

# ---- ambient accent glow (top-right) --------------------------------------
glow = Image.new("RGBA", (w, h), (0, 0, 0, 0))
gdr = ImageDraw.Draw(glow)
cx, cy, r = int(w * 0.86), int(h * 0.16), int(w * 0.34)
gdr.ellipse([cx - r, cy - r, cx + r, cy + r], fill=ACCENT_1 + (60,))
glow = glow.filter(ImageFilter.GaussianBlur(110 * S // 2))
img = Image.alpha_composite(img, glow)

# soft indigo glow bottom-left for depth
glow2 = Image.new("RGBA", (w, h), (0, 0, 0, 0))
g2 = ImageDraw.Draw(glow2)
cx2, cy2, r2 = int(w * 0.08), int(h * 0.92), int(w * 0.30)
g2.ellipse([cx2 - r2, cy2 - r2, cx2 + r2, cy2 + r2], fill=ACCENT_2 + (40,))
glow2 = glow2.filter(ImageFilter.GaussianBlur(120 * S // 2))
img = Image.alpha_composite(img, glow2)
draw = ImageDraw.Draw(img, "RGBA")


# ---- helpers for the brand mark -------------------------------------------
def rounded_box(d, box, radius, **kw):
    d.rounded_rectangle(box, radius=radius, **kw)


def sparkle_path(d, ox, oy, scale, fill):
    """4-point sparkle (matches index.html / icon.svg mark)."""
    # control polygon from icon.svg sparkle, scaled
    pts = []
    # build a smooth 4-point star using bezier-ish sampling
    n = 200
    R = 48 * scale
    inner = 0.30
    for i in range(n + 1):
        a = (i / n) * 2 * math.pi - math.pi / 2
        # pinch profile: sharp points, concave sides
        k = abs(math.cos(2 * a))
        rad = R * (inner + (1 - inner) * (k ** 1.6))
        pts.append((ox + rad * math.cos(a), oy + rad * math.sin(a)))
    d.polygon(pts, fill=fill)


# ---- brand mark: spreadsheet card + sparkle (drawn on its own layer) -------
mark = Image.new("RGBA", (w, h), (0, 0, 0, 0))
md = ImageDraw.Draw(mark, "RGBA")

mark_scale = 1.06 * S
# position the mark group; vertically centered against the text column
gx, gy = int(116 * S), int(176 * S)

# spreadsheet card outline
card_w, card_h = int(232 * mark_scale), int(272 * mark_scale)
rounded_box(md, [gx, gy, gx + card_w, gy + card_h], radius=int(22 * S),
            outline=CARD_STROKE, width=int(8 * S))

# header row (accent gradient)
hdr_box = (gx + int(12 * S), gy + int(12 * S),
           gx + card_w - int(12 * S), gy + int(12 * S) + int(46 * S))
hw = hdr_box[2] - hdr_box[0]
hh = hdr_box[3] - hdr_box[1]
hdr_grad = diag_grad_rgba((hw, hh), ACCENT_1, ACCENT_2)
hdr_mask = Image.new("L", (hw, hh), 0)
ImageDraw.Draw(hdr_mask).rounded_rectangle([0, 0, hw - 1, hh - 1],
                                           radius=int(11 * S), fill=255)
mark.paste(hdr_grad, (hdr_box[0], hdr_box[1]), hdr_mask)

# grid dividers
v1 = gx + int(80 * mark_scale)
v2 = gx + int(150 * mark_scale)
top_div = gy + int(64 * S)
md.line([(v1, top_div), (v1, gy + card_h - int(12 * S))], fill=CARD_STROKE, width=int(7 * S))
md.line([(v2, top_div), (v2, gy + card_h - int(12 * S))], fill=CARD_STROKE, width=int(7 * S))
for yy in (110, 164, 218):
    y = gy + int(yy * mark_scale)
    md.line([(gx + int(12 * S), y), (gx + card_w - int(12 * S), y)],
            fill=CARD_STROKE, width=int(7 * S))

# sparkle (overlapping the card's lower-right) with soft glow
sp = Image.new("RGBA", (w, h), (0, 0, 0, 0))
spd = ImageDraw.Draw(sp, "RGBA")
sp_ox = gx + card_w - int(18 * S)
sp_oy = gy + card_h - int(64 * S)
sparkle_path(spd, sp_ox, sp_oy, 1.55 * S, ACCENT_1 + (255,))
# small secondary sparkle
sparkle_path(spd, sp_ox + int(58 * S), sp_oy + int(46 * S), 0.70 * S, ACCENT_LT + (255,))
# glow behind sparkle
sp_glow = sp.filter(ImageFilter.GaussianBlur(18 * S // 2))
mark = Image.alpha_composite(mark, sp_glow)
# recolor sparkle fill with gradient overlay
sp_grad = diag_grad_rgba((w, h), ACCENT_1, ACCENT_2)
sp_alpha = sp.split()[3]
sp_colored = Image.new("RGBA", (w, h), (0, 0, 0, 0))
sp_colored.paste(sp_grad, (0, 0), sp_alpha)
mark = Image.alpha_composite(mark, sp_colored)

img = Image.alpha_composite(img, mark)
draw = ImageDraw.Draw(img, "RGBA")

# ---- text block ------------------------------------------------------------
def tracked_text(d, pos, text, fnt, fill, tracking):
    """Draw text with manual letter-spacing; returns total advance width."""
    x, y = pos
    x0 = x
    for ch in text:
        d.text((x, y), ch, font=fnt, fill=fill)
        x += d.textlength(ch, font=fnt) + tracking
    return x - x0


def gradient_text(text, fnt, c1, c2):
    """Return (RGBA layer, size) of `text` filled with a diagonal gradient."""
    tmp = Image.new("RGBA", (8, 8), (0, 0, 0, 0))
    bbox = ImageDraw.Draw(tmp).textbbox((0, 0), text, font=fnt)
    lw = bbox[2] - bbox[0] + 6 * S
    lh = bbox[3] + 12 * S
    layer = Image.new("RGBA", (lw, lh), (0, 0, 0, 0))
    ImageDraw.Draw(layer).text((-bbox[0], 0), text, font=fnt, fill=(255, 255, 255, 255))
    grad = diag_grad_rgba((lw, lh), c1, c2)
    out = Image.new("RGBA", (lw, lh), (0, 0, 0, 0))
    out.paste(grad, (0, 0), layer.split()[3])
    return out


brand_x = int(432 * S)

# eyebrow / kicker
f_kick = font("seguisb.ttf", 25 * S)
tracked_text(draw, (brand_x + int(3 * S), int(150 * S)),
             "PROMPT  →  SPREADSHEET", f_kick, ACCENT_LT, int(5 * S))

# Brand wordmark
f_brand = font("segoeuib.ttf", 60 * S)
brand_y = int(196 * S)
draw.text((brand_x, brand_y), "Sheet", font=f_brand, fill=TEXT)
sheet_w = draw.textlength("Sheet", font=f_brand)
g_out = gradient_text("Genie", f_brand, ACCENT_1, ACCENT_LT)
img.alpha_composite(g_out, (int(brand_x + sheet_w + 2 * S), brand_y))
draw = ImageDraw.Draw(img, "RGBA")

# Headline (the tagline as the hero)
f_head = font("segoeuib.ttf", 76 * S)
draw.text((brand_x, int(300 * S)), "Speak or type.", font=f_head, fill=TEXT)
l2_out = gradient_text("Get a spreadsheet.", f_head, ACCENT_LT, ACCENT_2)
img.alpha_composite(l2_out, (brand_x, int(392 * S)))
draw = ImageDraw.Draw(img, "RGBA")

# Sub / supporting line
f_sub = font("segoeui.ttf", 31 * S)
draw.text((brand_x, int(508 * S)),
          "Describe what you need — get a real Excel spreadsheet.",
          font=f_sub, fill=TEXT_DIM)

# ---- thin accent rule at the very bottom -----------------------------------
bar = diag_grad_rgba((w, int(8 * S)), ACCENT_1, ACCENT_2)
img.alpha_composite(bar, (0, h - int(8 * S)))

# ---- downsample & save -----------------------------------------------------
final = img.convert("RGB").resize((W, H), Image.LANCZOS)
out = "C:/Users/flori/Documents/sheet-genie/public/og-image.png"
final.save(out, "PNG", optimize=True)
print("wrote", out, final.size)
