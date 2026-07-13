#!/usr/bin/env python3
"""Slide 08 — RISK, SECURITY & GOVERNANCE. The risks executives worry about most,
each mitigated by design. MECE: accuracy · data security · governance/control."""
from pptx.util import Inches
import deck_kit as K


def render(prs):
    s = K.page("Risk, security & governance",
               "The risks executives worry about most (accuracy, data security "
               "and auditability) are mitigated by design, not bolted on")

    cards = [
        ("Accuracy & trust", K.NAVY, [
            "Answers come only from approved bank content",
            "Every answer shows its source, so it can be checked",
            "Expired offers are flagged; if there's no answer, it says so",
        ]),
        ("Data security & privacy", K.AMBER, [
            "Runs in a tightly controlled, secure environment",
            "Can run entirely inside the bank, no data leaves",
            "Never handles customer secrets such as PINs or passwords",
        ]),
        ("Content governance (CMS Hub)", K.GOLD, [
            "Product Owners keep the CMS Hub current, the single source of truth",
            "Start and end dates set at upload, so expired offers are never served",
            "Dual control on uploads; every change is logged and auditable",
        ]),
    ]

    y = K.TOP + Inches(0.12)
    ch = Inches(3.35)
    for (title, accent, points), (cx, cw) in zip(cards, K.cols(3)):
        K.rect(s, cx, y, cw, ch, K.CARD, line=K.RULE, lw=1.0)
        K.rect(s, cx, y, cw, Inches(0.085), accent)
        K.text(s, cx + Inches(0.24), y + Inches(0.26), cw - Inches(0.48), Inches(0.5),
               [[(title, 15.5, K.NAVY, True)]], sa=0)
        K.bullets(s, cx + Inches(0.24), y + Inches(0.86), cw - Inches(0.46),
                  points, size=12.5, gap=11)

    # bottom takeaway band — ties governance to the To-Be process
    K.callout(s, K.ML, Inches(5.75), K.CW, Inches(0.72), "Bottom line:  ",
              "the process assigns clear ownership: Product Owners must keep materials "
              "and their start/end dates current in the CMS Hub, so OSG always answers "
              "from approved, in-force content, with every answer traceable to source.")
    return s
