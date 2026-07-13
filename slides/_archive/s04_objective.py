"""Slide 4 — Objective (drivers on the left, success bar on the right)."""
import deck_kit as K


def render(prs):
    s = K.page(
        "Objective",
        "Equip every front-line employee with a trusted, self-updating assistant "
        "for the bank's offerings",
    )

    col = K.cols(5)               # 5-col grid → 60/40 split (3 : 2)
    lx, lw = K.spanx(col, 0, 2)   # left: columns 0-2
    rx, rw = K.spanx(col, 3, 4)   # right: columns 3-4

    # Left — the objective, unpacked.
    K.bullets(s, lx, K.TOP + K.Inches(0.12), lw, [
        ("Serve the people who serve customers",
         "relationship managers, tellers and customer service."),
        ("Answer promos, products & services",
         "what they are, terms, eligibility, and how to apply — bilingual (ID/EN)."),
        ("Make correctness the default",
         "quote figures exactly; when unsure, say so — never guess."),
        ("Stay current with zero maintenance burden",
         "the knowledge updates itself as source content changes."),
    ], size=14.5, gap=17)

    # Right — "what good looks like" card panel.
    cy = K.TOP
    ch = K.Inches(4.20)           # ends at ~6.10"
    K.rect(s, rx, cy, rw, ch, K.CARD, line=K.RULE, lw=1.0)
    K.rect(s, rx, cy, rw, K.Inches(0.085), K.TEAL)   # colored top rule
    K.text(s, rx + K.Inches(0.26), cy + K.Inches(0.30), rw - K.Inches(0.52), K.Inches(0.4),
           [[("What “good” looks like", 14, K.NAVY, True)]], sa=0)
    K.bullets(s, rx + K.Inches(0.26), cy + K.Inches(0.88), rw - K.Inches(0.52), [
        "Every answer is source-cited and verifiable",
        "No expired offer ever shown as active",
        "New content live within seconds, no restart",
        "One model runs selection and synthesis",
        "Runs unprivileged on a locked-down host",
    ], size=12.5, gap=14)

    return s
