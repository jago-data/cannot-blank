#!/usr/bin/env python3
"""Testimonial-card primitives for the user-feedback slides (s11 / s12)."""
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import deck_kit as K

AV_BG = RGBColor(0xE6, 0xEF, 0xEC)


def avatar(s, x, y, d, accent):
    """A simple person-silhouette placeholder (user picture) in a circle."""
    K.rect(s, x, y, d, d, AV_BG, shape=MSO_SHAPE.OVAL)
    hd = Emu(int(int(d) * 0.30))
    K.rect(s, Emu(int(x) + int(d) // 2 - int(hd) // 2), Emu(int(y) + int(int(d) * 0.16)),
           hd, hd, accent, shape=MSO_SHAPE.OVAL)
    sw = Emu(int(int(d) * 0.60))
    sh = Emu(int(int(d) * 0.46))
    K.rect(s, Emu(int(x) + int(d) // 2 - int(sw) // 2), Emu(int(y) + int(int(d) * 0.54)),
           sw, sh, accent, shape=MSO_SHAPE.OVAL)


def card(s, x, y, w, h, accent, statement, name, title):
    K.rect(s, x, y, w, h, K.WHITE, line=K.RULE_D, lw=1.0,
           shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    K.rect(s, x, y, Inches(0.08), h, accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    K.text(s, x + Inches(0.26), y + Inches(0.06), Inches(1.0), Inches(0.7),
           [[("“", 44, accent, True)]], sa=0)
    K.text(s, x + Inches(0.30), y + Inches(0.86), w - Inches(0.58), h - Inches(1.86),
           [[(statement, 12, K.SLATE, False, True)]], sa=0, ls=1.28)
    K.hline(s, x + Inches(0.30), y + h - Inches(0.94), w - Inches(0.58), K.RULE, 0.9)
    ad = Inches(0.62)
    avatar(s, x + Inches(0.30), y + h - Inches(0.80), ad, accent)
    K.text(s, x + Inches(1.06), y + h - Inches(0.82), w - Inches(1.3), Inches(0.7),
           [[(name, 12.5, K.NAVY, True)], [(title, 10, K.GREY, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=1, ls=1.05)


def wall(s, cards):
    """Lay out up to three testimonial cards across the content band."""
    n = len(cards)
    gap = Inches(0.4)
    w = Emu(int((int(K.CW) - (n - 1) * int(gap)) / n))
    y, h = Inches(2.05), Inches(4.55)
    cx = int(K.ML)
    for accent, statement, name, title in cards:
        card(s, Emu(cx), y, w, h, accent, statement, name, title)
        cx += int(w) + int(gap)
