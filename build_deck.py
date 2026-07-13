#!/usr/bin/env python3
"""OSG — Smart Guide: McKinsey-style executive deck (16:9).

Assembler. The visual design system lives in `deck_kit.py`; each slide lives in its own
module under `slides/` exposing `render(prs)` and adding exactly one slide. This file
calls them in order so slide numbering + the shared footer stay correct.

Rebuild:  /home/hartonosng/miniforge3/envs/py_env/bin/python build_deck.py
"""
import os
import deck_kit as K

from slides import (
    s01_cover, s02_agenda, s03_exec, s04_charter, s05_problem,
    s06_proof, s07_process, s07b_process, s08_risk, s09_costbenefit,
    s11_feedback, s12_feedback, s13a_demo, s13_demo, s14_thanks,
)

# C-suite structure (extends the CLAUDE.md brief):
# Cover · Contents · Executive Summary · Project Charter · Background/Problem/Solution ·
# Proof (built & live) · Cross-Functional Process Map · Cost & Benefit ·
# Risk/Security/Governance · User feedback · Demo
ORDER = [
    s01_cover, s02_agenda, s03_exec, s04_charter, s05_problem,
    s06_proof, s07_process, s07b_process, s09_costbenefit, s08_risk,
    s11_feedback, s12_feedback, s13a_demo, s13_demo, s14_thanks,
]


def main():
    prs = K.new_deck()
    for mod in ORDER:
        mod.render(prs)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "OSG-Smart-Guide.pptx")
    prs.save(out)
    print(f"Saved {out}  ({len(prs.slides._sldIdLst)} slides)")


if __name__ == "__main__":
    main()
