#!/usr/bin/env python3
"""Slide 10 — RECOMMENDATION & NEXT STEPS (dark). The decision ask + closing lockup."""
from pptx.util import Inches, Pt
import deck_kit as K

STEPS = [
    ("Approve a scoped pilot",
     "one product line and one branch team, to quantify time-saved, accuracy and adoption vs today."),
    ("Confirm executive sponsorship",
     "name the sponsor and cross-functional owners: business, technology & security, and knowledge."),
    ("Set the success criteria & scale gate",
     "agree the metrics and thresholds that green-light a bank-wide rollout."),
]


def render(prs):
    s = K.dark_page()

    K.text(s, K.ML, Inches(1.24), K.CW, Inches(0.34),
           [[("RECOMMENDATION  ·  NEXT  STEPS", 13, K.GOLD, True)]], sa=0)

    y = int(Inches(2.04))
    row_h = int(Inches(1.02))
    for head, sub in STEPS:
        K.text(s, K.ML, Inches(0.02) + y, Inches(0.4), Inches(0.5),
               [[("▪", 16, K.AMBER, True)]], sa=0)
        K.text(s, K.ML + Inches(0.42), y, K.CW - Inches(0.42), Inches(0.9),
               [[(head, 18, K.WHITE, True)],
                [(sub, 13.5, K.ONNAVY, False)]], sa=3, ls=1.08)
        y += row_h

    K.rect(s, K.ML, Inches(5.58), Inches(0.92), Pt(2.4), K.AMBER)
    K.text(s, K.ML, Inches(5.80), K.CW, Inches(0.9),
           [[("OSG Smart Guide", 22, K.WHITE, True)],
            [("Grounded · Secure · Bilingual · Built and ready to pilot", 13.5, K.GOLD, True)]], sa=6, ls=1.1)
    return s
