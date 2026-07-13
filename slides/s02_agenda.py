#!/usr/bin/env python3
"""Slide 02 — CONTENTS (light). Five main sections, single column with a one-line
descriptor. The deck is 10 slides; the contents lists the 5 sections they group into."""
from pptx.util import Inches, Emu
from pptx.enum.text import MSO_ANCHOR
import deck_kit as K

SECTIONS = [
    ("01", "Executive summary",
     "The problem, the AI solution, and the projected impact at a glance"),
    ("02", "Project charter",
     "Scope, strategic alignment, high-level timeline, and executive sponsorship"),
    ("03", "Background, problem & proposed solution",
     "The cost of inefficiency today, the AI-driven enabler, and the proof it is built"),
    ("04", "Cross-functional process map",
     "As-Is vs To-Be workflow, time saved, and risk mitigated by design"),
    ("05", "Cost & benefit analysis",
     "CAPEX vs OPEX, projected payback, and the recommended pilot"),
]


def _item(s, x, w, y, row_h, num, label, desc):
    K.text(s, x, y, Inches(1.05), row_h,
           [[(num, 26, K.AMBER, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, x + Inches(1.05), y, w - Inches(1.05), row_h,
           [[(label, 17, K.NAVY, True)],
            [(desc, 12.5, K.GREY, False)]], anchor=MSO_ANCHOR.MIDDLE, sa=3, ls=1.05)
    K.hline(s, x, y + row_h, w, K.RULE, 0.9)


def render(prs):
    s = K.page("Contents",
               "From the problem to the solution to the business impact")

    x, w = K.ML, K.CW
    row_h = Inches(0.92)
    gap = Inches(0.06)
    y0 = K.TOP + Inches(0.14)
    for i, (num, label, desc) in enumerate(SECTIONS):
        y = Emu(int(y0) + i * (int(row_h) + int(gap)))
        _item(s, x, w, y, row_h, num, label, desc)
    return s
