---
name: "conference-proceedings-latex"
description: "Use when working on Chinese conference proceedings or abstract-book LaTeX projects generated from Excel metadata, especially tasks involving generate_tex.js, main_full.tex, toc_content.tex, papers_content.tex, pdflatex compilation, corresponding-author email placement, author/affiliation formatting, table-of-contents leader dots, page numbers, or fixing LaTeX escaping issues in proceedings."
---

# Conference Proceedings LaTeX

## Core Workflow
1. Inspect the current project before editing: identify `generate_tex.js`, `main_full.tex`, `main.tex`, `toc_content.tex`, `papers_content.tex`, and the Excel source such as `摘要信息.xls`.
2. Prefer changing generation logic in `generate_tex.js` when output is generated; do not hand-edit generated files except for diagnosis.
3. Keep `main_full.tex` and `main.tex` macros synchronized when changing shared layout behavior.
4. After changes, run `node generate_tex.js`, then compile the full book twice with `pdflatex -interaction=nonstopmode -jobname=main_full_toc_style main_full.tex` unless the user requests a different job name.
5. Scan the log for blocking errors and visually inspect rendered PDF pages when layout matters.

## References
- For layout and data rules, read `references/layout-rules.md`.
- For validation commands and known failure checks, read `references/validation.md`.

## Operating Rules
- Use the repository's existing LaTeX macros and generator structure; avoid introducing a new build system unless explicitly requested.
- Treat `toc_content.tex` and `papers_content.tex` as generated artifacts; regenerate them from Excel/source data after generator edits.
- For PDF layout changes, render representative pages with `pdftoppm` and inspect PNGs before finalizing.
- Report both the generated PDF path and the verification performed.
