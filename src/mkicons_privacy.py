# -*- coding: utf-8 -*-
"""White line icons (159x159, transparent) in the style of the GT template icons: padlock, shield-with-check, check-in-circle.
Run once; the PNGs are written to src/assets."""
import os
from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
S = 4; N = 159 * S; W = 7 * S   # supersample 4x, stroke 7px


def canvas():
    im = Image.new("RGBA", (N, N), (0, 0, 0, 0)); return im, ImageDraw.Draw(im)


def done(im, name):
    im.resize((159, 159), Image.LANCZOS).save(os.path.join(OUT, name + ".png"))


def p(x, y): return (x * S, y * S)


# padlock
im, d = canvas()
d.rounded_rectangle([p(36, 70), p(123, 128)], radius=10 * S, outline="white", width=W)
d.arc([p(52, 26), p(107, 100)], 180, 360, fill="white", width=W)
d.line([p(52, 63), p(52, 72)], fill="white", width=W); d.line([p(107, 63), p(107, 72)], fill="white", width=W)
d.ellipse([p(73, 92), p(86, 105)], outline="white", width=W // 2 + 2)
d.line([p(79.5, 105), p(79.5, 116)], fill="white", width=W // 2 + 2)
done(im, "ic_lock")

# shield with check
im, d = canvas()
pts = [p(79.5, 22), p(124, 38), p(124, 80), p(79.5, 136), p(35, 80), p(35, 38)]
d.line(pts + [pts[0]], fill="white", width=W, joint="curve")
d.line([p(58, 78), p(74, 95), p(104, 60)], fill="white", width=W, joint="curve")
done(im, "ic_shield")

# check in circle
im, d = canvas()
d.ellipse([p(28, 28), p(131, 131)], outline="white", width=W)
d.line([p(54, 82), p(72, 100), p(106, 60)], fill="white", width=W, joint="curve")
done(im, "ic_check")
print("icons written to", OUT)
