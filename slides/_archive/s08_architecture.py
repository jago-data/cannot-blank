#!/usr/bin/env python3
"""Slide 08 — ARCHITECTURE. Three-box stack + pluggable-LLM band."""
from pptx.util import Inches
import deck_kit as K


def render(prs):
    s = K.page("How it works",
               "A thin, pluggable stack: React + FastAPI + DuckDB, with a "
               "swappable LLM and no vector infrastructure")

    boxes = [
        ("Frontend", "React 18 · Vite · Tailwind\nChat · Document View · Admin",
         K.TEAL),
        ("Backend API", "FastAPI · Uvicorn\nSSE streaming · LLM-Wiki", K.AMBER),
        ("Knowledge store", "DuckDB + Parquet\nMarkdown pages, no DB server",
         K.GOLD),
    ]

    by = K.TOP + Inches(0.10)
    bh = Inches(1.55)
    cspec = K.cols(3)
    for (title, body, accent), (cx, cw) in zip(boxes, cspec):
        K.rect(s, cx, by, cw, bh, K.CARD, line=K.RULE, lw=1.0)
        K.rect(s, cx, by, cw, Inches(0.085), accent)
        K.text(s, cx + Inches(0.24), by + Inches(0.28), cw - Inches(0.48),
               bh - Inches(0.4),
               [[(title, 14.5, K.NAVY, True)], [(body, 11, K.GREY, False)]],
               sa=6, ls=1.12)

    # left-right arrows joining the boxes, sitting in each gutter
    ay = by + (bh - Inches(0.42)) / 2
    for i in range(len(boxes) - 1):
        box_right = int(cspec[i][0]) + int(cspec[i][1])
        ax = box_right + (int(K.GUT) - int(Inches(0.22))) // 2
        K.rect(s, ax, ay, Inches(0.22), Inches(0.42), K.FAINT,
               shape=K.MSO_SHAPE.LEFT_RIGHT_ARROW)

    # ── pluggable-LLM band ──
    band_y = by + bh + Inches(0.35)
    band_h = Inches(1.75)
    K.rect(s, K.ML, band_y, K.CW, band_h, K.BAND)
    K.rect(s, K.ML, band_y, Inches(0.07), band_h, K.AMBER)
    K.text(s, K.ML + Inches(0.30), band_y + Inches(0.06),
           K.CW - Inches(0.6), band_h - Inches(0.12),
           [[("Pluggable LLM — one line in config.yaml", 14.5, K.NAVY, True)],
            [("anthropic (Claude Haiku, default)   ·   openai_compatible "
              "(remote Qwen / any /v1)   ·   llamacpp (local CPU GGUF)",
              12.5, K.INK, False)],
            [("Backend :8200  ·  Vite :5174 proxies /api (single origin → "
              "no CORS)  ·  ingest via pypdf + watchdog auto-reindex",
              12, K.GREY, False)]],
           anchor=K.MSO_ANCHOR.MIDDLE, sa=10, ls=1.1)
    return s
