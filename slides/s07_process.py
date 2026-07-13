#!/usr/bin/env python3
"""Slide 07 — CROSS-FUNCTIONAL PROCESS MAP · BEFORE (without OSG), swimlane flowchart."""
from pptx.util import Inches, Emu
import deck_kit as K
from slides import _swim as SW


def render(prs):
    s = K.page(
        "Cross-functional process map · before (without OSG)",
        "Today, a manual chain from Product Owner to branch, with an H+1 escalation "
        "whenever the answer is not on hand",
        title_size=19,
    )

    y0 = Inches(1.98)
    lane_h = Inches(1.60)
    _, _, (CUST, FRONT, PO) = SW.lanes(s, y0, lane_h, [
        ("CUSTOMER", K.NAVY, K.CARD),
        ("FRONTLINER", K.TEAL_D, K.CARD2),
        ("PRODUCT OWNER", K.GOLD, K.MINT),
    ])

    c = [int(Inches(v)) for v in (2.95, 5.05, 7.15, 9.25, 11.35)]

    b1 = SW.box(s, c[0], CUST, "Customer asks a question")
    b2 = SW.box(s, c[0], FRONT, "Search the materials the PO uploaded")
    d1 = SW.diamond(s, c[1], FRONT, "Material found?")
    byes = SW.box(s, c[2], FRONT, "Answer the customer, same day")
    bc1 = SW.box(s, c[2], CUST, "Customer helped, same day")
    bno = SW.box(s, c[1], PO, "Escalate to the PIC PO; PO replies (H+1)")
    bfin = SW.box(s, c[4], FRONT, "Follow up and answer the customer")
    bc2 = SW.box(s, c[4], CUST, "Customer helped, next day (H+1)")

    SW.connect(s, b1, b2)
    SW.connect(s, b2, d1)
    SW.connect(s, d1, byes, branch=("YES", K.TEAL))
    SW.connect(s, byes, bc1)
    SW.connect(s, d1, bno, branch=("NO", K.AMBER))
    SW.connect(s, bno, bfin)
    SW.connect(s, bfin, bc2)
    return s
