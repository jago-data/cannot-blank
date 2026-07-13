#!/usr/bin/env python3
"""Slide 07b — CROSS-FUNCTIONAL PROCESS MAP · AFTER (with OSG), swimlane flowchart."""
from pptx.util import Inches, Emu
import deck_kit as K
from slides import _swim as SW


def render(prs):
    s = K.page(
        "Cross-functional process map · after (with OSG)",
        "With OSG, one centralized CMS Hub feeds the AI, and the frontliner answers "
        "from the source in seconds",
        title_size=19,
    )

    y0 = Inches(1.96)
    lane_h = Inches(1.23)
    _, _, (CUST, FRONT, OSG, PO) = SW.lanes(s, y0, lane_h, [
        ("CUSTOMER", K.NAVY, K.CARD),
        ("FRONTLINER", K.TEAL_D, K.CARD2),
        ("OSG · AI", K.AMBER, K.MINT),
        ("CMS HUB · PO", K.GOLD, K.CARD),
    ])

    c = [int(Inches(v)) for v in (2.85, 4.95, 7.05, 9.15)]
    bh = Inches(0.62)

    n_up = SW.box(s, c[0], PO, "Upload materials to the CMS Hub", h=bh)
    n_pull = SW.box(s, c[0], OSG, "OSG pulls from CMS Hub; KB updates", h=bh)
    n_ask = SW.box(s, c[1], CUST, "Customer asks a question", h=bh)
    n_log = SW.box(s, c[1], FRONT, "Log in to OSG and search", h=bh)
    d = SW.diamond(s, c[2], OSG, "Answer found?", h=Inches(0.94))
    n_ans = SW.box(s, c[3], OSG, "Answer with source and full detail", h=bh)
    n_give = SW.box(s, c[3], FRONT, "Answer the customer, in seconds", h=bh)
    n_help = SW.box(s, c[3], CUST, "Customer helped, in seconds", h=bh)
    n_upd = SW.box(s, c[2], PO, "PIC PO updates the CMS Hub", h=bh)

    SW.connect(s, n_up, n_pull)
    SW.connect(s, n_pull, d)
    SW.connect(s, n_ask, n_log)
    SW.connect(s, n_log, d)
    SW.connect(s, d, n_ans, branch=("YES", K.TEAL))
    SW.connect(s, n_ans, n_give)
    SW.connect(s, n_give, n_help)
    SW.connect(s, d, n_upd, branch=("NO", K.AMBER))
    return s
