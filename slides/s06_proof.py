#!/usr/bin/env python3
"""Slide 06 — PROOF, BUILT & LIVE TODAY. Real live-pilot usage (6 Apr to 6 Jul 2026):
1,313 users, 12,944 questions, 86% answered correctly, plus the question mix and the
honest gaps behind the 14% it cannot yet answer.
"""
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import deck_kit as K


def _hbar(s, x, y, w, label, pct, color):
    K.text(s, x, y, Inches(2.3), Inches(0.28),
           [[(label, 11, K.NAVY, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)
    tx = x + Inches(2.4)
    tw = Emu(int(w) - int(Inches(2.4)) - int(Inches(0.55)))
    K.rect(s, tx, y + Inches(0.04), tw, Inches(0.20), K.RULE)          # track
    K.rect(s, tx, y + Inches(0.04), Emu(int(int(tw) * pct / 100.0)), Inches(0.20), color)
    K.text(s, Emu(int(tx) + int(tw) + int(Inches(0.06))), y, Inches(0.5), Inches(0.28),
           [[("%d%%" % pct, 11, color, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)


def render(prs):
    s = K.page("Proof · built & live today",
               "Live since April 2026: 1,313 users have asked 12,944 questions, with "
               "86% answered correctly",
               source="Source: OSG live-pilot usage, 6 Apr to 6 Jul 2026")

    # ── KPI strip ─────────────────────────────────────────────────────────
    strip_y = K.TOP + Inches(0.04)
    strip_h = Inches(1.30)
    K.rect(s, K.ML, strip_y, K.CW, strip_h, K.CARD, line=K.RULE, lw=1.0)
    kpis = [
        ("1,313", "users", K.AMBER),
        ("12,944", "questions asked", K.TEAL),
        ("326", "products & promos live", K.GOLD),
        ("86%", "answered correctly", K.TEAL_D),
    ]
    colk = K.cols(4)
    for (num, lab, col), (cx, cw) in zip(kpis, colk):
        K.stat_tile(s, cx, strip_y + Inches(0.20), cw, num, lab, accent=col, num_size=30)
    for i in range(1, 4):
        vx = colk[i][0] - int(K.GUT) // 2
        K.vline(s, K.Emu(vx), strip_y + Inches(0.24), strip_h - Inches(0.48), K.RULE_D, 1.0)
    K.text(s, K.ML, strip_y + strip_h + Inches(0.06), K.CW, Inches(0.24),
           [[("Live pilot · 6 Apr to 6 Jul 2026 · adoption and accuracy climbing toward "
              "the 95% target", 10.5, K.GREY, False)]], sa=0)

    # ── lower band: question mix (left) + answer quality & gaps (right) ────
    ly = strip_y + strip_h + Inches(0.46)
    (lx, lw), (rx, rw) = K.cols(2)

    K.text(s, lx, ly, lw, Inches(0.28),
           [[("WHAT USERS ASK", 11, K.NAVY, True)]], sa=0)
    _hbar(s, lx, ly + Inches(0.42), lw, "Product & promo", 76, K.AMBER)
    _hbar(s, lx, ly + Inches(0.86), lw, "Operational terms & policy", 15, K.GOLD)
    _hbar(s, lx, ly + Inches(1.30), lw, "Greetings", 8, K.TEAL)
    K.text(s, lx, ly + Inches(1.86), lw, Inches(0.6),
           [[("Most questions are exactly what OSG is built for, product and promo "
              "answers, and it handles them in seconds.", 10.5, K.SLATE, False)]],
           sa=0, ls=1.2)

    # right: answer quality split + honest gaps
    K.text(s, rx, ly, rw, Inches(0.28),
           [[("ANSWER QUALITY", 11, K.NAVY, True)]], sa=0)
    bar_y = ly + Inches(0.42)
    bw86 = Emu(int(int(rw) * 0.86))
    K.rect(s, rx, bar_y, bw86, Inches(0.32), K.TEAL_D)
    K.rect(s, Emu(int(rx) + int(bw86)), bar_y, Emu(int(rw) - int(bw86)), Inches(0.32), K.AMBER)
    K.text(s, rx + Inches(0.12), bar_y, bw86, Inches(0.32),
           [[("86% answered correctly", 11, K.WHITE, True)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, Emu(int(rx) + int(bw86)), bar_y, Emu(int(rw) - int(bw86)), Inches(0.32),
           [[("14%", 11, K.WHITE, True)]], align=PP_ALIGN.CENTER,
           anchor=MSO_ANCHOR.MIDDLE, sa=0)
    # big highlighted 14% callout — the deliberate "won't guess" gap
    hl_y = bar_y + Inches(0.46)
    hl_h = Inches(0.72)
    K.rect(s, rx, hl_y, rw, hl_h, K.MINT)
    K.rect(s, rx, hl_y, Inches(0.07), hl_h, K.AMBER)
    K.text(s, rx + Inches(0.20), hl_y, Inches(1.4), hl_h,
           [[("14%", 33, K.AMBER, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, rx + Inches(1.66), hl_y, rw - Inches(1.82), hl_h,
           [[("are not wrong answers, they are questions OSG deliberately does not "
              "guess on:", 11, K.NAVY, True)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=0, ls=1.14)
    K.bullets(s, rx, bar_y + Inches(1.34), rw, [
        "Internal data or policy questions",
        "Internal procedures and authorization",
        "Real-time data, such as currency rates",
        "Data outside the current period, e.g. a future-month promo",
    ], size=10.5, gap=6, marker=K.AMBER)
    return s
