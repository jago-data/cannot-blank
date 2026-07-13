#!/usr/bin/env python3
"""Slide 07 — DEMO. A real staff Q&A: mock chat exchange (left) + what it shows (right)."""
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import deck_kit as K


def render(prs):
    s = K.page("Demo",
               "In practice: a staff question returns a concise, source-cited "
               "answer in seconds")

    c = K.cols(3)
    lx, lw = K.spanx(c, 0, 1)          # chat exchange spans columns 1–2
    rx, rw = c[2]                      # "what it shows" panel on column 3

    # ── user question bubble (right-aligned feel) ──
    ubw = Inches(4.7)
    ubx = lx + (lw - ubw)
    uby = K.TOP
    K.rect(s, ubx, uby, ubw, Inches(0.6), K.TEAL,
           shape=K.MSO_SHAPE.ROUNDED_RECTANGLE)
    K.text(s, ubx + Inches(0.22), uby, ubw - Inches(0.44), Inches(0.6),
           [[("“limit cash collateral loan berapa?”", 13, K.WHITE, True)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=0)

    # ── assistant answer bubble ──
    aby = uby + Inches(0.80)
    abh = Inches(2.02)
    abw = Inches(7.0)
    K.rect(s, lx, aby, abw, abh, K.WHITE, line=K.RULE, lw=1.0,
           shape=K.MSO_SHAPE.ROUNDED_RECTANGLE)
    K.text(s, lx + Inches(0.26), aby + Inches(0.06), abw - Inches(0.52),
           abh - Inches(0.12),
           [[("Limit Cash Collateral Loan adalah maksimum Rp25 Miliar ",
              12.5, K.INK, False), ("[1]", 10, K.AMBER, True),
             (".", 12.5, K.INK, False)],
            [("Pinjaman dapat diajukan mulai dari Rp50 Juta hingga Rp25 Miliar, "
              "dengan jaminan deposito atau aset cair lainnya ",
              12.5, K.INK, False), ("[1]", 10, K.AMBER, True),
             (".", 12.5, K.INK, False)],
            [("Mau aku jelaskan syarat atau jangka waktu pinjaman ini?",
              12, K.GREY, False)]],
           anchor=MSO_ANCHOR.MIDDLE, sa=8, ls=1.08)

    # ── source chip ──
    scy = aby + abh + Inches(0.22)
    K.chip(s, lx, scy, Inches(3.3), Inches(0.5), "[1]  Cash Collateral Loan",
           fill=K.CARD, line=K.RULE, txt=K.NAVY, size=11.5, bold=True)

    # ── caption ──
    K.text(s, lx, scy + Inches(0.62), abw, Inches(0.35),
           [[("→ click the source to open the Document View, scoped to "
              "that page.", 11, K.GREY, False)]], sa=0)

    # ── right panel: what the demo shows ──
    K.text(s, rx, aby - Inches(0.02), rw, Inches(0.4),
           [[("The demo shows", 14, K.NAVY, True)]], sa=0)
    K.bullets(s, rx, aby + Inches(0.50), rw, [
        "Answer-first, numbered, concise",
        "Exact figure quoted from source",
        "[n] citation on every claim",
        "Follow-up offered",
        "One click → the source document",
    ], size=12.5, gap=13)
    return s
