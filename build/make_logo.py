# -*- coding: utf-8 -*-
"""Готує два варіанти фірмового логотипа OBRIO для презентації.

Джерело — оригінальний файл із деку користувача (ppt/media/image1.png),
тобто справжній брендовий асет, а не відтворення.

- logo_brand.png — оригінальні кольори, для темних слайдів;
- logo_ink.png   — той самий контур у кольорі тексту деку, для світлих слайдів
                   (бірюза на білому дає контраст близько 1,7:1 — замало).
"""
import zipfile
from collections import Counter
from pathlib import Path

from PIL import Image

BUILD = Path(__file__).resolve().parent
ASSETS = BUILD / "assets"
ASSETS.mkdir(exist_ok=True)
SOURCE_DECK = BUILD.parent.parent / "submission" / "01_OBRIO_creative_testing_framework.pptx"

INK = (23, 22, 53)      # C.ink деку
SCALE = 4               # ресемплінг, щоб PowerPoint не масштабував піксельний вихідник

with zipfile.ZipFile(SOURCE_DECK) as z:
    media = [n for n in z.namelist() if n.startswith("ppt/media/")]
    assert len(media) == 1, f"очікували один медіафайл, знайшли {media}"
    raw = z.read(media[0])
(ASSETS / "logo_source.png").write_bytes(raw)

img = Image.open(ASSETS / "logo_source.png").convert("RGBA")
w, h = img.size
pixels = img.load()

visible = [(r, g, b) for x in range(w) for y in range(h)
           for r, g, b, a in [pixels[x, y]] if a > 200]
top = Counter(visible).most_common(3)
opaque = sum(1 for x in range(w) for y in range(h) if pixels[x, y][3] > 10)
print(f"вихідник: {w}×{h}, непрозорих пікселів {opaque} ({opaque / (w * h):.0%})")
print("основні кольори:", [("#%02X%02X%02X" % c, n) for c, n in top])


def relative_luminance(rgb):
    def channel(value):
        value /= 255
        return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4
    r, g, b = (channel(v) for v in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a, b):
    la, lb = relative_luminance(a), relative_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


brand = top[0][0]
print(f"контраст бренду до білого:      {contrast(brand, (255, 255, 255)):.2f}:1")
print(f"контраст бренду до темного тла:  {contrast(brand, INK):.2f}:1")
print(f"контраст ink-версії до білого:   {contrast(INK, (255, 255, 255)):.2f}:1")

big = (w * SCALE, h * SCALE)

# 1. Брендовий варіант для темних слайдів — кольори не чіпаємо.
img.resize(big, Image.LANCZOS).save(ASSETS / "logo_brand.png")

# 2. Монохромний варіант для світлих слайдів: зберігаємо альфу, замінюємо RGB.
ink = Image.new("RGBA", img.size, INK + (0,))
ink.putalpha(img.getchannel("A"))
ink.resize(big, Image.LANCZOS).save(ASSETS / "logo_ink.png")

print(f"збережено: assets/logo_brand.png і assets/logo_ink.png ({big[0]}×{big[1]})")
print(f"пропорції: {w / h:.3f}")
