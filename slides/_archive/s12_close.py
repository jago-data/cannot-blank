#!/usr/bin/env python3
"""Slide 12 — CLOSE (dark). Recommended next steps + closing lockup."""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import deck_kit as K

STEPS = [
    ("Run a scoped pilot",
     "one product line and one branch team; measure time-saved and accuracy vs today."),
    ("Prioritize the hard-facts capture",
     "fees, rates and eligibility that sit behind JavaScript tabs on the source site."),
    ("Ship the catalog revamp",
     "structured browse-by-category so staff self-serve beyond chat."),
]


def render(prs):
    s = K.dark_page()

    # kicker
    K.text(s, K.ML, Inches(1.28), K.CW, Inches(0.34),
           [[("RECOMMENDED  NEXT  STEPS", 12, K.GOLD, True)]], sa=0)

    # bulleted list — bold white head + muted subtext, amber/gold markers
    y = int(Inches(2.10))
    row_h = int(Inches(1.02))
    for head, sub in STEPS:
        K.text(s, K.ML, Inches(0.02) + y, Inches(0.4), Inches(0.5),
               [[("▪", 15, K.AMBER, True)]], sa=0)
        K.text(s, K.ML + Inches(0.42), y, K.CW - Inches(0.42), Inches(0.9),
               [[(head, 18, K.WHITE, True)],
                [(sub, 12.5, K.ONNAVY, False)]], sa=3, ls=1.08)
        y += row_h

    # amber tick divider
    K.rect(s, K.ML, Inches(5.62), Inches(0.92), Pt(2.4), K.AMBER)

    # closing lockup
    K.text(s, K.ML, Inches(5.84), K.CW, Inches(0.9),
           [[("OSG — Smart Guide", 22, K.WHITE, True)],
            [("Grounded · Pluggable · No RAG · Bilingual · Docker-free deploy",
              12.5, K.GOLD, True)]], sa=6, ls=1.1)
    return s
