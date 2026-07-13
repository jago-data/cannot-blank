#!/usr/bin/env python3
"""Slide 10 — IMPACT. Three benefit cards (Faster / Safer / Cheaper) + caption."""
from pptx.util import Inches
import deck_kit as K


def render(prs):
    s = K.page("Impact",
               "OSG is designed to cut lookup time and remove fabrication "
               "risk while staying maintainable")

    cards = [
        ("Faster",
         "Seconds to a precise, cited answer instead of searching multiple "
         "portals and PDFs — staff stay with the customer.", K.TEAL),
        ("Safer",
         "Every figure is grounded and cited; expired offers are flagged; "
         "customer secrets (OTP/PIN/CVV) are refused by design.", K.AMBER),
        ("Cheaper to run & scale",
         "No embeddings/vector DB and a pluggable model keep cost low; "
         "self-updating knowledge keeps maintenance near zero.", K.GOLD),
    ]

    y = K.TOP + Inches(0.35)
    ch = Inches(2.75)
    for (title, body, accent), (cx, cw) in zip(cards, K.cols(3)):
        K.card(s, cx, y, cw, ch, title, body, accent=accent,
               title_size=16, body_size=12.5)

    cy = y + ch + Inches(0.42)
    K.text(s, K.ML, cy, K.CW, Inches(0.5),
           [[("Qualitative today; a pilot would quantify time-saved-per-query "
              "and answer-accuracy against the current process.",
              11.5, K.GREY, False)]], sa=0)

    return s
