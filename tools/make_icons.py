"""Turn a real potsticker photo into glowing PWA icons.

usage: python3 make_icons.py <photo> <out_dir> [--focus x,y] [--zoom z]
  --focus  centre of the potsticker in the photo, as fractions (default 0.5,0.5)
  --zoom   how tightly to crop around it (1 = largest centred square)
"""
import sys, math, random, argparse
from PIL import Image, ImageDraw, ImageFilter, ImageChops, ImageEnhance

ap = argparse.ArgumentParser()
ap.add_argument('photo'); ap.add_argument('out')
ap.add_argument('--focus', default='0.5,0.5'); ap.add_argument('--zoom', type=float, default=1.0)
a = ap.parse_args()
fx, fy = map(float, a.focus.split(','))

src = Image.open(a.photo).convert('RGB')
W, H = src.size
side = min(W, H) / a.zoom
cx, cy = fx * W, fy * H
l = max(0, min(W - side, cx - side / 2)); t = max(0, min(H - side, cy - side / 2))
photo = src.crop((round(l), round(t), round(l + side), round(t + side)))

S = 1024  # master size, downscaled at the end


def radial(size, stops):
    """stops: list of (pos 0..1, (r,g,b,a)) from centre outwards."""
    img = Image.new('RGBA', (size, size))
    px = img.load(); c = (size - 1) / 2
    for y in range(size):
        for x in range(size):
            d = min(1, math.hypot(x - c, y - c) / c)
            for (p0, c0), (p1, c1) in zip(stops, stops[1:]):
                if p0 <= d <= p1:
                    f = (d - p0) / (p1 - p0 or 1)
                    px[x, y] = tuple(round(c0[i] + (c1[i] - c0[i]) * f) for i in range(4)); break
    return img


def build(scale, full_bleed):
    base = radial(S // 4, [(0, (92, 58, 32, 255)), (.55, (52, 34, 22, 255)), (1, (22, 15, 10, 255))]).resize((S, S), Image.BICUBIC)
    # golden halo behind the food
    halo = radial(S // 4, [(0, (255, 205, 120, 255)), (.55, (240, 170, 85, 220)), (.8, (210, 130, 55, 90)), (1, (200, 120, 50, 0))]).resize((S, S), Image.BICUBIC)
    hs = int(S * min(1.15, scale * 1.65)); halo = halo.resize((hs, hs), Image.BICUBIC)
    base.alpha_composite(halo, ((S - hs) // 2, (S - hs) // 2))

    # the photo in a soft-edged circle
    ps = int(S * scale)
    p = photo.resize((ps, ps), Image.LANCZOS)
    p = ImageChops.multiply(p, Image.new('RGB', p.size, (255, 226, 180))); p = ImageEnhance.Brightness(p).enhance(1.12); p = ImageEnhance.Color(p).enhance(1.25); p = ImageEnhance.Contrast(p).enhance(1.06)
    edge = radial(ps // 4, [(0, (0, 0, 0, 255)), (.55, (0, 0, 0, 255)), (.85, (235, 160, 80, 255)), (1, (250, 190, 110, 255))]).resize((ps, ps), Image.BICUBIC).convert('RGB')
    p = ImageChops.screen(p, edge)
    mask = Image.new('L', (ps, ps), 0)
    inset = int(ps * .1)
    ImageDraw.Draw(mask).ellipse((inset, inset, ps - inset, ps - inset), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(ps * .05))
    layer = Image.new('RGBA', (S, S)); off = ((S - ps) // 2, (S - ps) // 2)
    layer.paste(p.convert('RGBA'), off, mask)
    base.alpha_composite(layer)

    # bloom: brightest parts of the food glow outward
    rgb = base.convert('RGB')
    lum = rgb.convert('L').point(lambda v: max(0, v - 135) * 1.9)
    bright = Image.composite(rgb, Image.new('RGB', (S, S)), lum)
    warm = ImageChops.multiply(bright, Image.new('RGB', (S, S), (255, 214, 150)))
    for r, k in ((S * .01, .6), (S * .04, .8), (S * .1, .8)):
        g = warm.filter(ImageFilter.GaussianBlur(r))
        rgb = ImageChops.screen(rgb, ImageEnhance.Brightness(g).enhance(k))

    # rim light around the circle
    rim = Image.new('L', (S, S), 0); d = ImageDraw.Draw(rim)
    r0 = ps / 2 - inset; c = S / 2
    d.ellipse((c - r0, c - r0, c + r0, c + r0), outline=255, width=max(2, int(S * .006)))
    rim = rim.filter(ImageFilter.GaussianBlur(S * .01))
    rgb = ImageChops.screen(rgb, Image.composite(Image.new('RGB', (S, S), (255, 200, 120)), Image.new('RGB', (S, S)), rim))

    # a few sparkles
    rnd = random.Random(7); sp = Image.new('L', (S, S), 0); d = ImageDraw.Draw(sp)
    for ang, dist, sz in ((-50, .62, .05), (35, .66, .035), (150, .6, .03), (-140, .64, .025)):
        x = c + math.cos(math.radians(ang)) * ps * dist; y = c + math.sin(math.radians(ang)) * ps * dist
        L = S * sz; w = L * .16
        d.polygon([(x, y - L), (x + w, y), (x, y + L), (x - w, y)], fill=255)
        d.polygon([(x - L, y), (x, y + w), (x + L, y), (x, y - w)], fill=255)
    sp = ImageChops.lighter(sp.filter(ImageFilter.GaussianBlur(S * .004)), sp.filter(ImageFilter.GaussianBlur(S * .012)))
    rgb = ImageChops.screen(rgb, Image.composite(Image.new('RGB', (S, S), (255, 236, 200)), Image.new('RGB', (S, S)), sp))

    img = rgb.convert('RGBA')
    if not full_bleed:  # rounded-square shape for normal icons
        m = Image.new('L', (S, S), 0)
        ImageDraw.Draw(m).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * .22), fill=255)
        img.putalpha(m)
    return img


icon = build(.68, False)          # standard icon
mask = build(.58, True)           # maskable: food inside the 80% safe zone, full bleed
apple = build(.70, True)          # iOS rounds corners itself and dislikes transparency

import os
os.makedirs(a.out, exist_ok=True)
for img, name, size in ((icon, 'icon-512.png', 512), (icon, 'icon-192.png', 192), (mask, 'icon-maskable-512.png', 512),
                        (apple, 'apple-touch-icon.png', 180), (icon, 'favicon-32.png', 32)):
    img.resize((size, size), Image.LANCZOS).save(os.path.join(a.out, name), optimize=True)
icon.resize((512, 512), Image.LANCZOS).save(os.path.join(a.out, 'preview-icon.png'))
print('ok')
