#!/usr/bin/env python3
"""Slide 05 — BACKGROUND, PROBLEM & SOLUTION (executive dashboard, 3-act).

A linear story told as three colour-coded act bands — 01 BACKGROUND, 02 PROBLEM,
03 SOLUTION — each with a left act-rail and a row of premium metric tiles (pain
metrics under Problem, outcome metrics under Solution). Charter figures woven in.
Design selected from a 5-designer study (variant v5).
"""
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import deck_kit as K

MINT_BG = RGBColor(0xFB, 0xF1, 0xF2)   # faint red wash — the PROBLEM act
TEAL_BG = RGBColor(0xED, 0xF4, 0xF2)   # faint steel wash — the SOLUTION act
NEUT_BG = K.CARD                        # neutral — the BACKGROUND act


def _in(v):
    return Inches(v)


NUM_X, NUM_W = _in(0.82), _in(0.66)
NAME_X, NAME_W = _in(1.50), _in(1.60)
ZONE_X = _in(3.30)
ZONE_W = Emu(int(_in(0.75 + 11.833)) - int(ZONE_X))


def act_band(s, y, h, tint, spine):
    K.rect(s, K.ML, y, K.CW, h, tint, line=K.RULE, lw=1.0)
    K.rect(s, K.ML, y, _in(0.075), h, spine)


def act_rail(s, y, h, num, name, desc, accent):
    K.text(s, NUM_X, y, NUM_W, h, [[(num, 30, accent, True)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, NAME_X, y, NAME_W, h,
           [[(name, 13, K.NAVY, True)], [(desc, 9.5, K.GREY, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=2, ls=1.04)


def tiles(s, zone_x, zone_w, y, h, items, accent, n=4, gap=_in(0.18)):
    tw = Emu(int((int(zone_w) - (n - 1) * int(gap)) / n))
    for i, (num, unit, label, sub) in enumerate(items):
        x = Emu(int(zone_x) + i * (int(tw) + int(gap)))
        K.rect(s, x, y, tw, h, K.WHITE, line=K.RULE, lw=1.0,
               shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        K.rect(s, x, y, tw, _in(0.06), accent)   # top accent rule
        px = x + _in(0.18)
        pw = Emu(int(tw) - int(_in(0.36)))
        K.text(s, px, y + _in(0.15), pw, _in(0.44),
               [[(num, 23, accent, True), (unit, 11, accent, True)]], sa=0, wrap=False)
        K.text(s, px, y + _in(0.64), pw, _in(0.28),
               [[(label, 11, K.NAVY, True)]], sa=0, ls=1.02)
        K.text(s, px, y + _in(0.92), pw, _in(0.40),
               [[(sub, 9, K.GREY, False)]], sa=0, ls=1.02)


def render(prs):
    s = K.page(
        "Background, problem & solution",
        "About 6,000 customer questions a day take ~49,300 hours to answer today; OSG "
        "returns a cited answer in seconds, ~8 hours in total",
        source="Source: internal analysis · EES 2025, NPS H1 2025",
        title_size=19,
    )

    # ═════════════ ACT 01 · BACKGROUND ═════════════
    A1Y, A1H = _in(1.92), _in(0.92)
    act_band(s, A1Y, A1H, NEUT_BG, K.TEAL_D)
    act_rail(s, A1Y, A1H, "01", "BACKGROUND", "The bank's front line", K.TEAL_D)
    K.text(s, ZONE_X + _in(0.05), A1Y, _in(2.7), A1H,
           [[("1,200", 30, K.NAVY, True)],
            [("frontliners, the bank's front line", 10.5, K.SLATE, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=2, ls=1.02)
    K.vline(s, _in(6.15), A1Y + _in(0.20), Emu(int(A1H) - int(_in(0.40))), K.RULE_D, 1.0)
    K.text(s, _in(6.40), A1Y, _in(2.9), A1H,
           [[("~6,000", 30, K.NAVY, True)],
            [("customer questions every day", 10.5, K.SLATE, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=2, ls=1.02)
    K.vline(s, _in(9.55), A1Y + _in(0.20), Emu(int(A1H) - int(_in(0.40))), K.RULE_D, 1.0)
    K.text(s, _in(9.80), A1Y, _in(2.7), A1H,
           [[("Every question needs a fast, correct and trusted answer, on first "
              "contact.", 11, K.SLATE, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=0, ls=1.10)

    # ═════════════ ACT 02 · PROBLEM ═════════════
    A2Y, A2H = _in(2.99), _in(1.74)
    act_band(s, A2Y, A2H, MINT_BG, K.AMBER)
    act_rail(s, A2Y, A2H, "02", "PROBLEM", "Scattered knowledge, slow answers", K.AMBER)
    tiles(s, ZONE_X, ZONE_W, A2Y + _in(0.20), _in(1.34), [
        ("493", " min", "elapsed per case", "escalate to a Product Owner, wait to H+1"),
        ("49,300", " hrs", "per day, all frontliners", "493 min x 1,200 x 5 cases"),
        ("36%", "", "cannot find the answer", "email, intranet, old PDFs · EES 2025"),
        ("38%", "", "of customers complain", "late, inconsistent answers · NPS H1 2025"),
    ], K.AMBER)

    # transformation cue between the two acts
    K.arrow(s, _in(1.63), _in(4.60), _in(0.32), _in(0.26), color=K.TEAL_D,
            shape=MSO_SHAPE.DOWN_ARROW)

    # ═════════════ ACT 03 · SOLUTION ═════════════
    A3Y, A3H = _in(4.88), _in(1.74)
    act_band(s, A3Y, A3H, TEAL_BG, K.TEAL_D)
    act_rail(s, A3Y, A3H, "03", "SOLUTION", "OSG: one AI source of truth", K.TEAL_D)
    tiles(s, ZONE_X, ZONE_W, A3Y + _in(0.20), _in(1.34), [
        ("~5", " sec", "cited answer per case", "ask in plain words, get the source"),
        ("~8", " hrs", "per day, all frontliners", "5 sec x 1,200 x 5 cases"),
        ("100%", "", "source-cited & current", "one approved single source of truth"),
        ("1st", "-contact", "resolved & traceable", "confident, correct, auditable"),
    ], K.TEAL_D)
    return s
