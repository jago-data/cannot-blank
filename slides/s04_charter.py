#!/usr/bin/env python3
"""Slide 04 — PROJECT CHARTER (light, enhanced UI).

Four titled cards with filled headers (Situational context, Objective & key results,
Problem statement, Scope & constraints) over a high-level timeline strip with the real
schedule (CMS upload through launch). Real figures from EES 2025 / NPS H1 2025 render
as text; nothing in the timeline is bold.
"""
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import deck_kit as K

HDR = Inches(0.42)
PAD = Inches(0.26)


def _card(s, px, py, pw, ph, title):
    """White card with a graphite header band + white title. Returns (ix, iy, iw)."""
    K.rect(s, px, py, pw, ph, K.WHITE, line=K.RULE_D, lw=1.0)
    K.rect(s, px, py, pw, HDR, K.NAVY)
    K.text(s, px + PAD, py, pw - 2 * PAD, HDR,
           [[(title, 13.5, K.WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)
    return px + PAD, py + HDR + Inches(0.18), pw - 2 * PAD


def _badge(s, x, y, w, h, label, fill, txt=K.WHITE, size=10.5, bold=True):
    K.rect(s, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    K.text(s, x, y, w, h, [[(label, size, txt, bold)]],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)


def render(prs):
    s = K.page(
        "Project charter",
        "One AI-based single source of truth, targeting ≤ 5-second lookups, "
        "95% answer accuracy and cutting customer complaints to < 10%",
        title_size=20,
    )

    (lx, pw), (rx, _) = K.cols(2)
    r1y, r1h = Inches(1.80), Inches(2.08)
    r2y, r2h = Inches(4.02), Inches(1.86)

    # ── TL — Situational context (key figures emphasized in red) ──────────
    ix, iy, iw = _card(s, lx, r1y, pw, r1h, "Situational context")
    K.text(s, ix, iy, iw, Inches(1.5), [
        [("▪  ", 10.5, K.AMBER, True),
         ("EES 2025:  ", 10.5, K.NAVY, True), ("36%", 10.5, K.AMBER, True),
         (" of frontliners (", 10.5, K.SLATE, False), ("1,200+", 10.5, K.AMBER, True),
         (") cannot easily find current product, promo and procedure information; it is "
          "scattered across email, intranet and other channels, and mixed with expired "
          "material.", 10.5, K.SLATE, False)],
        [("▪  ", 10.5, K.AMBER, True),
         ("No single source of truth", 10.5, K.NAVY, True),
         (", frontliners search manually or ask Product Owners directly, who then field "
          "hundreds to thousands of repeat questions, cutting productivity at branches "
          "and Head Office.", 10.5, K.SLATE, False)],
    ], sa=10, ls=1.14)

    # ── TR — Objective & key results (metric badges) ──────────────────────
    ix, iy, iw = _card(s, rx, r1y, pw, r1h, "Objective & key results")
    K.text(s, ix, iy, iw, Inches(0.26),
           [[("Objective   ", 10.5, K.NAVY, True),
             ("one AI single source of truth for product & promo information.",
              11, K.GREY, False)]], sa=0)
    krs = [
        ("≤ 5 sec", K.AMBER, "frontliner lookup (from >30 min)"),
        ("95%",     K.TEAL,  "accuracy of frontliner answers"),
        ("< 10%",   K.GOLD,  "NPS: weak frontliner advisory"),
        ("↓",       K.NAVY,  "frontliner reliance on POs"),
    ]
    bw, bh = Inches(0.98), Inches(0.28)
    ky = iy + Inches(0.32)
    for i, (metric, col, desc) in enumerate(krs):
        y = Emu(int(ky) + i * int(Inches(0.27)))
        K.text(s, ix, y, Inches(0.5), bh, [[("KR%d" % (i + 1), 10, K.FAINT, True)]],
               anchor=MSO_ANCHOR.MIDDLE, sa=0)
        _badge(s, ix + Inches(0.48), y, bw, bh, metric, col)
        descx = Emu(int(ix) + int(Inches(0.48)) + int(bw) + int(Inches(0.14)))
        descw = Emu(int(iw) - int(Inches(0.48)) - int(bw) - int(Inches(0.14)))
        K.text(s, descx, y, descw, bh,
               [[(desc, 10.5, K.SLATE, False)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)

    # ── BL — Problem statement (colored category tags) ────────────────────
    ix, iy, iw = _card(s, lx, r2y, pw, r2h, "Problem statement")
    probs = [
        ("Productivity", K.TEAL,
         "frontliners wait up to next-day (H+1) for answers, from manual search and "
         "Head-Office bottlenecks."),
        ("Accuracy", K.GOLD,
         "answers to customers vary and can be biased, with no single source of truth."),
        ("Experience", K.AMBER,
         "NPS H1 2025:  38% of customers complained about frontliner answer quality."),
    ]
    for i, (tag, col, desc) in enumerate(probs):
        y = Emu(int(iy) + i * int(Inches(0.44)))
        K.text(s, ix, y, iw, Inches(0.42),
               [[("● ", 10, col, True), (tag + ":  ", 10.5, K.NAVY, True),
                 (desc, 10.5, K.GREY, False)]], sa=0, ls=1.12)

    # ── BR — Scope (pill chips) & constraints ─────────────────────────────
    ix, iy, iw = _card(s, rx, r2y, pw, r2h, "Scope & constraints")
    K.text(s, ix, iy, iw, Inches(0.4),
           [[("Scope   ", 10.5, K.NAVY, True),
             ("Central platform · Content management · Training · Monitoring",
              10.5, K.SLATE, False)]], sa=0, ls=1.12)
    K.hline(s, ix, iy + Inches(0.40), iw, K.RULE, 0.8)
    K.bullets(s, ix, iy + Inches(0.50), iw, [
        "PO uploads verified before publishing.",
        "Dual control on every upload to the content hub.",
        "Start/end dates prevent use of expired content.",
    ], size=10.5, gap=6, marker=K.NAVY)

    # ── Bottom — high-level timeline (real schedule; nothing bold) ─────────
    K.text(s, K.ML, Inches(6.04), K.CW, Inches(0.22),
           [[("HIGH-LEVEL TIMELINE", 10.5, K.SLATE, False),
             ("   ·  content upload from Nov 2025, live by Apr 2026",
              10.5, K.GREY, False)]], sa=0)
    phases = [
        ("PO content upload to CMS Hub", "W4 Nov 2025 to W2 Feb 2026", K.TEAL),
        ("UAT: OSG Admin & SQM", "W3 Feb 2026", K.TEAL_D),
        ("UAT: OSG by Product Owners", "W3 Mar 2026", K.GOLD),
        ("Launch OSG", "W2 Apr 2026", K.AMBER),
    ]
    by, bh2 = Inches(6.28), Inches(0.48)
    n = len(phases)
    pgap = Inches(0.26)
    bw2 = Emu(int((int(K.CW) - (n - 1) * int(pgap)) / n))
    for i, (title, date, col) in enumerate(phases):
        bx = Emu(int(K.ML) + i * (int(bw2) + int(pgap)))
        K.rect(s, bx, by, bw2, bh2, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        K.text(s, Emu(int(bx) + int(Inches(0.10))), by,
               Emu(int(bw2) - int(Inches(0.20))), bh2,
               [[(title, 10.5, K.WHITE, False)],
                [(date, 9.5, K.RGBColor(0xEB, 0xEC, 0xEE), False)]],
               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=1, ls=1.02)
        if i < n - 1:
            K.arrow(s, Emu(int(bx) + int(bw2) + int(pgap) // 2 - int(Inches(0.09))),
                    Emu(int(by) + int(bh2) // 2 - int(Inches(0.10))),
                    Inches(0.18), Inches(0.22), color=K.FAINT)
    return s
