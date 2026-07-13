#!/usr/bin/env python3
"""Slide 06 — APPROACH. Four-step process flow + why this beats classic RAG."""
from pptx.util import Inches
import deck_kit as K


def render(prs):
    s = K.page("Approach",
               "OSG compiles knowledge into a maintained wiki and lets the "
               "LLM itself retrieve it — no RAG, no embeddings")

    steps = [
        ("Sources", "Cleaned Markdown from OCBC + PDFs", K.TEAL),
        ("Wiki", "Category hubs, index, health, graph", K.GOLD),
        ("LLM retrieval", "Two stages: pick category → pick page", K.AMBER),
        ("Cited answer", "Grounded, with [n] citations", K.NAVY),
    ]
    K.chevron_flow(s, K.ML, Inches(2.15), K.CW, Inches(1.5), steps)

    K.text(s, K.ML, Inches(4.05), K.CW, Inches(0.4),
           [[("Why this beats classic RAG", 15, K.NAVY, True)]], sa=0)

    K.bullets(s, K.ML, Inches(4.62), K.CW, [
        ("Knowledge is compiled once and kept current",
         "not re-derived from raw chunks on every question — cross-references "
         "and summaries persist."),
        ("Fully LLM-driven, one model",
         "the same chat model routes categories then pages; no second model, "
         "no embeddings, no vector DB."),
        ("Accuracy guardrails built in",
         "figures cite the original source; synthesized prose is navigational "
         "only; expired offers are flagged, and “nothing relevant” is "
         "respected."),
    ], size=13, gap=10)
    return s
