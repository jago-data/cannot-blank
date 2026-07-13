# OSG — Smart Guide · Executive deck

`OSG-Smart-Guide.pptx` — a McKinsey-style **C-suite** briefing (16:9, 10 slides) for bank
leadership, generated from code. Follows the brief in `CLAUDE.md`: action titles, top-down
(Minto), MECE, low visual complexity. Palette is the brand **spec red `#E4002B`** as the single
accent over graphite/steel/ochre neutrals.

## Structure (10-slide C-suite storyline)
1. **Cover**
2. **Contents**
3. **Executive summary** — problem · AI solution · projected impact, in one snapshot
4. **Project charter** — scope, strategic alignment, timeline, executive sponsorship
5. **Background, problem & proposed solution** — cost of inaction vs the AI-driven enabler (SCQA)
6. **Proof — built & live today** — the platform already exists (2,758 docs · 51 hubs) + composition chart
7. **Cross-functional process map** — As-Is vs To-Be swimlanes, time saved across the value chain
8. **Risk, security & governance** — accuracy · data security · auditability, by design (MECE)
9. **Cost & benefit analysis** — CAPEX vs OPEX, payback, efficiency / revenue / risk (MECE)
10. **Recommendation & next steps** — the pilot ask and the path to scale

## Architecture of the build
- `deck_kit.py` — the **design system** (single source of visual truth): palette, 12-column
  grid, type scale, and a component library (`page`, `card`, `bullets`, `callout`, `chevron_flow`,
  `column_chart`, `stat_tile`, `lane`/`step`/`arrow`, and the fill-in `data_slot`/`kpi_slot`).
- `slides/sNN_*.py` — one module per slide, each exposing `render(prs)` and composing from the kit.
- `build_deck.py` — the **assembler**: calls the slide modules in order and saves the `.pptx`.
- `slides/_archive/` — the earlier 12-slide narrative deck (Objective→Roadmap), kept for reference.
- `build_deck_legacy.py` — the original single-file build, kept for reference.

## Fill-in figures (important)
Every number that is **not** sourced from the repo (payback, CAPEX/OPEX, minutes saved, FTE
hours, adoption targets, sponsor names) renders as a **dashed amber `data_slot`** — a visible
"insert value here" box. Replace those with real figures before presenting. Only these are real,
drawn from the codebase / knowledge base: **2,758 documents · 379 product pages · 2,379 promo
pages · 51 category hubs**, bilingual (ID/EN), lean/no-vector-DB, pluggable LLM.

## Rebuild
```bash
/home/hartonosng/miniforge3/envs/py_env/bin/python build_deck.py
```
Requires `python-pptx` (already in the `py_env` conda env). Edit `deck_kit.py` (design) or a
`slides/sNN_*.py` (content), then re-run to regenerate `OSG-Smart-Guide.pptx` in place.

## Preview (optional, Windows PowerPoint)
`_export.ps1` drives PowerPoint via COM to export each slide to PNGs under `_preview/`:
```bash
powershell.exe -ExecutionPolicy Bypass -File "C:\\Users\\user\\Documents\\ai\\osg-prod\\ppt\\_export.ps1"
```
