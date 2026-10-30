"""App icon designs for Policy Review Snapshot (drawn at 4x and scaled down for smooth edges).

draw_icon(style, size, pad) -> PIL.Image.  Styles:
  report  - report sheet with coverage bars and a tick
  meter   - coverage ring (gauge) with a shield in the middle
  shield  - shield with a magnifying glass (review)
  bao     - the character 保 on a soft badge
"""
from PIL import Image, ImageDraw, ImageFont

PURPLE, PLUM, PINK, CORAL = (0x6B, 0x3F, 0xA0), (0x43, 0x16, 0x5F), (0xD9, 0x4F, 0x8C), (0xFF, 0x8A, 0x7A)
WHITE, LILAC = (255, 255, 255), (0xF3, 0xEC, 0xFA)


def _gradient(n, a, b, diagonal=True):
    img = Image.new("RGB", (n, n))
    px = img.load()
    for y in range(n):
        for x in range(n):
            t = ((x + y) / (2 * n)) if diagonal else y / n
            px[x, y] = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return img


def _report(d, n, s):
    # paper sheet
    x0, y0, x1, y1 = n * (0.25 + s), n * (0.18 + s), n * (0.75 - s), n * (0.82 - s)
    w = x1 - x0
    d.rounded_rectangle([x0, y0, x1, y1], radius=w * 0.09, fill=WHITE)
    # folded corner
    c = w * 0.22
    d.polygon([(x1 - c, y0), (x1, y0 + c), (x1 - c, y0 + c)], fill=LILAC)
    # header line
    d.rounded_rectangle([x0 + w * 0.14, y0 + w * 0.16, x0 + w * 0.6, y0 + w * 0.25], radius=w * 0.04, fill=PURPLE)
    # coverage bars (track + fill)
    for i, frac in enumerate((0.9, 0.55, 0.75)):
        by = y0 + w * (0.42 + i * 0.2)
        d.rounded_rectangle([x0 + w * 0.14, by, x1 - w * 0.14, by + w * 0.1], radius=w * 0.05, fill=LILAC)
        d.rounded_rectangle([x0 + w * 0.14, by, x0 + w * 0.14 + (w * 0.72) * frac, by + w * 0.1], radius=w * 0.05,
                            fill=PINK if i == 1 else PURPLE)
    # tick badge
    r = w * 0.2
    cx, cy = x1 - w * 0.02, y1 - w * 0.02
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=PINK, outline=WHITE, width=int(w * 0.04))
    d.line([(cx - r * 0.45, cy), (cx - r * 0.1, cy + r * 0.35), (cx + r * 0.5, cy - r * 0.35)], fill=WHITE,
           width=int(r * 0.28), joint="curve")


def _meter(d, n, s):
    cx = cy = n / 2
    R = n * (0.31 - s)
    t = n * 0.075
    d.ellipse([cx - R, cy - R, cx + R, cy + R], outline=(255, 255, 255, 90), width=int(t))
    d.arc([cx - R, cy - R, cx + R, cy + R], start=-90, end=200, fill=WHITE, width=int(t))
    # small shield
    w = R * 0.95
    top = cy - w * 0.55
    pts = [(cx, top), (cx + w / 2, top + w * 0.16), (cx + w / 2, top + w * 0.55), (cx, top + w * 1.05),
           (cx - w / 2, top + w * 0.55), (cx - w / 2, top + w * 0.16)]
    d.polygon(pts, fill=WHITE)
    d.line([(cx - w * 0.2, top + w * 0.5), (cx - w * 0.03, top + w * 0.67), (cx + w * 0.24, top + w * 0.33)],
           fill=PINK, width=int(w * 0.1), joint="curve")


def _shield(d, n, s):
    m = n * (0.2 + s)
    w = n - 2 * m
    cx = n / 2 - w * 0.06
    top = m
    pts = [(cx, top), (cx + w * 0.42, top + w * 0.14), (cx + w * 0.42, top + w * 0.5), (cx, top + w * 0.95),
           (cx - w * 0.42, top + w * 0.5), (cx - w * 0.42, top + w * 0.14)]
    d.polygon(pts, fill=WHITE)
    # magnifier over the shield
    r = w * 0.2
    gx, gy = cx + w * 0.06, top + w * 0.45
    d.ellipse([gx - r, gy - r, gx + r, gy + r], outline=PURPLE, width=int(w * 0.07))
    d.line([(gx + r * 0.7, gy + r * 0.7), (gx + r * 1.9, gy + r * 1.9)], fill=PINK, width=int(w * 0.11))
    # small bars inside the lens
    for i, h in enumerate((0.35, 0.65, 0.5)):
        bx = gx - r * 0.55 + i * r * 0.4
        d.rectangle([bx, gy + r * 0.45 - r * h, bx + r * 0.25, gy + r * 0.45], fill=PINK if i == 1 else PURPLE)


def _bao(d, n, s, font_path="msyhbd.ttc"):
    R = n * (0.33 - s)
    cx = cy = n / 2
    d.ellipse([cx - R, cy - R, cx + R, cy + R], fill=WHITE)
    try:
        f = ImageFont.truetype(font_path, int(R * 1.25))
    except OSError:
        f = ImageFont.truetype("msyh.ttc", int(R * 1.25))
    d.text((cx, cy + R * 0.04), "保", font=f, fill=PURPLE, anchor="mm")
    # small check accent
    r = R * 0.3
    bx, by = cx + R * 0.72, cy + R * 0.72
    d.ellipse([bx - r, by - r, bx + r, by + r], fill=PINK, outline=WHITE, width=int(R * 0.06))
    d.line([(bx - r * 0.45, by), (bx - r * 0.1, by + r * 0.35), (bx + r * 0.5, by - r * 0.35)], fill=WHITE,
           width=int(r * 0.3), joint="curve")


STYLES = {"report": (_report, PURPLE, PINK), "meter": (_meter, PLUM, PINK),
          "shield": (_shield, PURPLE, CORAL), "bao": (_bao, PURPLE, PINK)}


def draw_icon(style, size, pad=0.0):
    """pad > 0 shrinks the artwork (for 'maskable' icons that Android crops to a circle)."""
    fn, a, b = STYLES[style]
    n = size * 4
    img = _gradient(n // 4, a, b).resize((n, n))
    over = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    fn(ImageDraw.Draw(over), n, pad)
    img = Image.alpha_composite(img.convert("RGBA"), over)
    return img.resize((size, size), Image.LANCZOS).convert("RGB")
