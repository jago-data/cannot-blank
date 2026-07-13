#!/usr/bin/env python3
"""Slide 03 — EXECUTIVE SUMMARY (light).

Top-down / Minto: a governing-thought band, then a Problem → Solution → Prize
rule-of-three, then an impact-snapshot strip. Every projected/financial figure is
a dashed fill-in (kpi_slot); only the repo-real 2,758 documents render as a stat.
"""
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
import deck_kit as K


def render(prs):
    s = K.page(
        "Executive summary",
        "A secure, self-updating AI assistant can cut frontliner lookup time and "
        "error risk today, with a low-cost, low-risk path to bank-wide rollout",
    )

    # 1) Governing-thought band — the bottom-line, spanning full content width.
    K.callout(
        s, K.ML, K.TOP, K.CW, Inches(0.82), None,
        "OSG lets frontliners ask in plain language and returns a source-cited answer in "
        "seconds, grounded only in the bank's own content, so speed no longer trades off "
        "against accuracy or compliance.",
    )

    # 2) Problem → Solution → Prize (rule of three).
    cy = Inches(2.92)
    ch = Inches(2.08)
    pillars = [
        ("The problem",
         "Knowledge is scattered across portals and PDFs and changes constantly; "
         "manual lookup is slow, and a wrong or expired figure is a compliance and "
         "trust risk.", K.NAVY),
        ("The solution",
         "A secure AI assistant that answers only from the bank's own approved "
         "content and shows the source for every figure. Expired offers are flagged. "
         "Works in Bahasa Indonesia and English.", K.TEAL),
        ("The prize",
         "Faster service and higher first-contact resolution, fewer errors, and new "
         "cross-sell moments, at a low running cost.",
         K.AMBER),
    ]
    for (cx, cw), (title, body, accent) in zip(K.cols(3), pillars):
        K.card(s, cx, cy, cw, ch, title, body, accent=accent)

    # 3) Impact-snapshot strip.
    K.text(s, K.ML, Inches(5.14), K.CW, Inches(0.3),
           [[("Impact snapshot · projected from the cost-benefit case, with live proof "
              "in gold", 12, K.SLATE, True)]],
           sa=0)

    strip_y = Inches(5.52)
    slots = K.cols(4)
    K.stat_tile(s, slots[0][0], strip_y, slots[0][1], "~5 sec",
                "answer per query, vs up to 493 min today", accent=K.TEAL, num_size=30)
    K.stat_tile(s, slots[1][0], strip_y, slots[1][1], "< 3 mo",
                "projected payback on ~Rp 206 Mio a year", accent=K.AMBER, num_size=30)
    K.stat_tile(s, slots[2][0], strip_y, slots[2][1], "~Rp 4.1 Bn",
                "estimated Year-1 net benefit", accent=K.GOLD, num_size=26)
    # REAL proof metric from the repo.
    K.stat_tile(s, slots[3][0], strip_y, slots[3][1], "2,758",
                "documents already live · 51 categories", accent=K.GOLD, num_size=30)

    return s
