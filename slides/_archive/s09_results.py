#!/usr/bin/env python3
"""Slide 09 — RESULTS TO DATE. The deck's key data slide: a KPI strip, a
composition column chart on the left, and navigation/trust bullets on the right."""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import deck_kit as K


def render(prs):
    s = K.page("Results to date",
               "2,758 real OCBC documents, cleaned and organized into 51 "
               "navigable hubs — only in-force offers served")

    # ── (a) KPI strip ─────────────────────────────────────────────────────
    strip_y = K.TOP + Inches(0.06)
    strip_h = Inches(1.34)
    K.rect(s, K.ML, strip_y, K.CW, strip_h, K.CARD, line=K.RULE, lw=1.0)

    kpis = [
        ("2,758", "documents ingested", K.AMBER, None),
        ("379",   "product pages",      K.TEAL,  None),
        ("2,379", "promo pages",        K.GOLD,  None),
        ("51",    "category hubs",      K.NAVY,  "37 products + 14 promos"),
    ]
    colk = K.cols(4)
    ty = strip_y + Inches(0.20)
    for (num, lab, accent, sub), (cx, cw) in zip(kpis, colk):
        K.stat_tile(s, cx, ty, cw, num, lab, accent=accent, sub=sub, num_size=30)
    # thin hairlines between the four tiles
    for i in range(1, 4):
        vx = colk[i][0] - int(K.GUT) // 2
        K.vline(s, K.Emu(vx), strip_y + Inches(0.26), strip_h - Inches(0.52),
                K.RULE_D, 1.0)

    # ── lower band: chart (left) + bullets (right) ────────────────────────
    ly = strip_y + strip_h + Inches(0.34)          # ≈ 3.64"
    (lx, lw), (rx, rw) = K.cols(2)

    # (b) column chart — the 2,758 composition
    K.text(s, lx, ly, lw, Inches(0.3),
           [[("The 2,758 documents, by type", 13, K.NAVY, True)]], sa=0)
    K.legend(s, lx, ly + Inches(0.34),
             [("Product pages", K.TEAL), ("Promo pages", K.GOLD)])
    K.column_chart(s, lx, ly + Inches(0.72), lw, Inches(2.35),
                   series=[("Product pages", 379, K.TEAL),
                           ("Promo pages", 2379, K.GOLD)])

    # (c) navigation / trust bullets on the right
    K.bullets(s, rx, ly + Inches(0.04), rw, [
        ("Sourced & cleaned from the public OCBC site",
         "chrome, boilerplate and duplicate “other-product” carousels "
         "stripped into focused KB pages."),
        ("Structured for navigation",
         "products by line (Individu, Korporasi, SME, Syariah…); promos into "
         "14 themes (Dining, Hotel & Travel, Special Event…)."),
        ("Trustworthy by design",
         "real validity windows gate what's served; a health self-check flags "
         "duplicates, stale and orphan pages."),
    ], size=13.5, gap=16)

    return s
