#!/usr/bin/env python3
"""Slide 09 — COST & BENEFIT ANALYSIS.

Full CAPEX/OPEX breakdown with quantity, price and Year 1 to Year 5 (Rp Mio), a total
per year row, and a benefits + 5-year-total + payback summary strip beneath.
"""
from pptx.util import Inches, Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import deck_kit as K


def render(prs):
    s = K.page(
        "Cost & benefit analysis",
        "A lean ~Rp 206 Mio a year to run, set against an estimated productivity and "
        "service benefit many times larger",
        title_size=19,
    )

    # ── column geometry for the cost table ──
    xi = K.ML
    xq = Inches(4.30)                                   # quantity (centre)
    wq = Inches(0.62)
    xp = Inches(4.98)                                   # price (centre)
    wp = Inches(1.30)
    wy = Inches(1.02)
    xy = [Emu(int(Inches(6.42)) + i * int(Inches(1.06))) for i in range(5)]
    w_item = Emu(int(xq) - int(xi) - int(Inches(0.12)))

    hy = Inches(1.90)
    hh = Inches(0.36)
    rh = Inches(0.278)

    def row_y(i):
        return Emu(int(hy) + int(hh) + Inches(1) // 25 + i * int(rh))

    # header row
    K.rect(s, K.ML, hy, K.CW, hh, K.NAVY)
    K.text(s, xi + Inches(0.16), hy, w_item, hh,
           [[("Cost item", 11, K.WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, xq, hy, wq, hh, [[("Qty", 10, K.WHITE, True)]],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)
    K.text(s, xp, hy, wp, hh, [[("Price (Mio)", 10, K.WHITE, True)]],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)
    for i, x in enumerate(xy):
        K.text(s, x, hy, wy, hh, [[("Y%d" % (i + 1), 10, K.WHITE, True)]],
               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)

    def subhead_row(i, label, accent):
        y = row_y(i)
        K.rect(s, K.ML, y, K.CW, rh, K.MINT)
        K.rect(s, K.ML, y, Inches(0.06), rh, accent)
        K.text(s, xi + Inches(0.16), y, w_item, rh, [[(label, 10, K.NAVY, True)]],
               anchor=MSO_ANCHOR.MIDDLE, sa=0)

    def item_row(i, item, qty, price, years, bold=False, tint=None, vcol=K.INK):
        y = row_y(i)
        if tint:
            K.rect(s, K.ML, y, K.CW, rh, tint)
        K.text(s, xi + Inches(0.28), y, w_item, rh,
               [[(item, 10, (K.NAVY if bold else K.INK), bold)]],
               anchor=MSO_ANCHOR.MIDDLE, sa=0)
        K.text(s, xq, y, wq, rh, [[(qty, 10, K.SLATE, False)]],
               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)
        K.text(s, xp, y, wp, rh, [[(price, 10, K.SLATE, False)]],
               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)
        for j, x in enumerate(xy):
            v = years[j] if j < len(years) else ""
            K.text(s, x, y, wy, rh, [[(v, 10, vcol, bold)]],
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, sa=0)

    subhead_row(0, "CAPEX  ·  one-time", K.TEAL)
    item_row(1, "Database (500 GB)", "1", "61.8", ["62"])
    subhead_row(2, "OPEX  ·  build & recurring", K.GOLD)
    item_row(3, "Database IT direct charge", "3", "3.75", ["11"])
    item_row(4, "Man-days, develop OSG", "30", "3.75", ["113"])
    item_row(5, "Man-days, IT deployment", "15", "3.75", ["56"])
    item_row(6, "Cloudera license", "1", "5,000", ["200", "200", "200", "200", "200"])
    item_row(7, "Cloudera maintenance", "1", "152", ["6", "6", "6", "6", "6"])
    K.hline(s, K.ML, row_y(8), K.CW, K.RULE_D, 1.0)
    item_row(8, "Total cost per year (Rp Mio)", "", "",
             ["448", "206", "206", "206", "206"], bold=True, tint=K.CARD2, vcol=K.AMBER)

    subhead_row(9, "POTENTIAL BENEFIT  ·  estimate (ramps with adoption)", K.TEAL_D)
    item_row(10, "Productivity & operations (time saved)", "", "",
             ["3,500", "5,600", "7,000", "7,000", "7,000"], vcol=K.GOLD)
    item_row(11, "Potential revenue (cross-sell & retention)", "", "",
             ["1,000", "1,800", "2,500", "2,500", "2,500"], vcol=K.GOLD)
    K.hline(s, K.ML, row_y(12), K.CW, K.RULE_D, 1.0)
    item_row(12, "Net per year (benefit less cost)", "", "",
             ["4,052", "7,194", "9,294", "9,294", "9,294"], bold=True, tint=K.MINT,
             vcol=K.TEAL_D)

    K.text(s, K.ML, Emu(int(row_y(13)) + int(Inches(0.03))), K.CW, Inches(0.5),
           [[("Costs are actual estimates (Rp Mio = million); license and maintenance "
              "priced as total contract value, with the annual charge in Y1 to Y5. "
              "5-year cost ~Rp 1,272 Mio.", 8.5, K.FAINT, False, True)],
            [("About 10 minutes saved per frontliner per day "
              "(1,200 frontliners, ~240 days) valued at ~Rp 0.15 Mio/hour and ramping "
              "with adoption, plus a conservative customer-experience and cross-sell "
              "uplift. To be validated in pilot.", 8.5, K.FAINT, False, True)]],
           sa=2, ls=1.18)
    return s
