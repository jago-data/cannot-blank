#!/usr/bin/env python3
"""Slide 11 — ROADMAP. Now / Next / Later phase columns with colored header
strips, plus a thin timeline arrow beneath."""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import deck_kit as K


def render(prs):
    s = K.page("Roadmap",
               "Next: richer synthesis, deeper facts, and an even friendlier "
               "catalog")

    phases = [
        ("Now", "Live",
         "Cleaned KB, category hubs, two-stage LLM retrieval, cited chat, "
         "admin dashboard.", K.TEAL),
        ("Next", "Weeks",
         "LLM-authored entity & comparison pages with cross-links; Document "
         "View catalog revamp.", K.AMBER),
        ("Later", "Quarter",
         "Capture hard facts behind JS tabs (fees/rates/eligibility); "
         "retrieval & graph performance; pilot metrics.", K.GOLD),
    ]

    y = K.TOP + Inches(0.22)
    ch = Inches(2.75)
    hdr = Inches(0.74)
    for (ph, when, body, accent), (cx, cw) in zip(phases, K.cols(3)):
        K.rect(s, cx, y, cw, ch, K.CARD, line=K.RULE, lw=1.0)
        K.rect(s, cx, y, cw, hdr, accent)
        K.text(s, cx + Inches(0.24), y, cw - Inches(0.48), hdr,
               [[(ph, 17, K.WHITE, True)],
                [(when.upper(), 10.5, K.WHITE, False)]],
               anchor=MSO_ANCHOR.MIDDLE, sa=1, ls=1.0)
        K.text(s, cx + Inches(0.24), y + hdr + Inches(0.26),
               cw - Inches(0.48), ch - hdr - Inches(0.42),
               [[(body, 12.5, K.GREY, False)]], sa=0, ls=1.14)

    # ── thin timeline arrow beneath ───────────────────────────────────────
    ay = y + ch + Inches(0.34)                     # ≈ 6.15"
    aw = int(K.CW) - int(Inches(0.5))
    K.hline(s, K.ML, ay, K.Emu(aw), K.RULE_D, 2.2)
    ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                            K.Emu(int(K.ML) + aw), ay - Inches(0.09),
                            Inches(0.42), Inches(0.20))
    ar.fill.solid(); ar.fill.fore_color.rgb = K.RULE_D
    ar.line.fill.background(); ar.shadow.inherit = False
    K.text(s, K.ML, ay + Inches(0.14), K.CW, Inches(0.3),
           [[("Increasing depth & maturity", 10.5, K.FAINT, False)]], sa=0)

    return s
