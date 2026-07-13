#!/usr/bin/env python3
"""Slide 14 — CLOSING. Soft-green gradient bookend to the cover: a simple thank-you
with the one-line value proposition and an invitation to discuss.
"""
from pptx.util import Inches, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import deck_kit as K

GT = RGBColor(0x2A, 0x74, 0x64)
GB = RGBColor(0x08, 0x2C, 0x27)
ONGREEN = RGBColor(0xD8, 0xEC, 0xE8)


def _grad(s, x, y, w, h, c1, c2, angle=60):
    sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.line.fill.background(); sp.shadow.inherit = False
    sp.fill.gradient()
    try:
        sp.fill.gradient_angle = angle
    except Exception:
        pass
    st = sp.fill.gradient_stops
    st[0].color.rgb = c1; st[0].position = 0.0
    st[1].color.rgb = c2; st[1].position = 1.0
    return sp


def render(prs):
    s = K.blank()
    _grad(s, 0, 0, K.W, K.H, GT, GB, angle=60)
    K.rect(s, 0, 0, K.W, Inches(0.10), K.TEAL)

    K.text(s, Inches(0), Inches(2.45), K.W, Inches(0.34),
           [[("OSG  ·  SMART GUIDE", 13, K.TEAL, True)]], align=PP_ALIGN.CENTER, sa=0)
    K.text(s, Inches(0), Inches(2.86), K.W, Inches(1.2),
           [[("Thank you", 54, K.WHITE, True)]], align=PP_ALIGN.CENTER, sa=0)
    K.rect(s, Emu(int(K.W) // 2 - int(Inches(0.55))), Inches(4.18), Inches(1.10),
           Inches(0.07), K.TEAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    K.text(s, Inches(0), Inches(4.48), K.W, Inches(0.8),
           [[("Faster, source-cited answers for every frontliner. "
              "Questions and discussion welcome.", 15, ONGREEN, False)]],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0, ls=1.2)
    return s
