#!/usr/bin/env python3
"""Slide 05 — CHALLENGES. Four structural challenges as lettered cards + implication."""
from pptx.util import Inches
import deck_kit as K


def render(prs):
    s = K.page("Challenges",
               "Four structural challenges make manual lookup slow, "
               "error-prone and risky")

    cards = [
        ("A", "Scattered knowledge",
         "Hundreds of product pages and thousands of time-limited promos, "
         "across portals and PDFs.", K.AMBER),
        ("B", "Constant change",
         "Promos start and expire continuously; a stale answer is worse "
         "than none.", K.TEAL),
        ("C", "High cost of error",
         "A wrong rate or an expired offer quoted to a customer is a "
         "compliance & trust risk.", K.GOLD),
        ("D", "Chatbots hallucinate",
         "Generic LLMs invent figures and can't cite a source — "
         "unacceptable in banking.", K.NAVY),
    ]

    y = K.TOP + Inches(0.08)
    ch = Inches(2.9)
    for (num, title, body, accent), (cx, cw) in zip(cards, K.cols(4)):
        K.card(s, cx, y, cw, ch, title, body, num=num, accent=accent)

    # implication band spanning full content width, anchored lower for balance
    cy = Inches(5.55)
    K.callout(s, K.ML, cy, K.CW, Inches(0.66),
              "Implication:  ",
              "staff need answers that are fast AND provably correct AND "
              "always up to date — at the same time.")
    return s
