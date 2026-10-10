# Validation

## Standard Regeneration And Compile
Run from the proceedings project root:

```powershell
node generate_tex.js
pdflatex -interaction=nonstopmode -jobname=main_full_toc_style main_full.tex
pdflatex -interaction=nonstopmode -jobname=main_full_toc_style main_full.tex
```

## Blocking Log Checks
Search the generated log:

```powershell
rg -n "^! |Fatal error|Misplaced alignment tab character|Undefined control sequence" main_full_toc_style.log
```

No matches means these blocking errors were not found. A nonzero `rg` exit code is expected when there are no matches.

## Visual Checks
Render representative TOC pages after layout changes:

```powershell
New-Item -ItemType Directory -Force tmp\pdfs | Out-Null
pdftoppm -png -f 2 -l 4 -r 160 main_full_toc_style.pdf tmp\pdfs\main_full_toc_style_check
```

Inspect the PNGs and verify:
- TOC title and section headings are readable.
- Paper IDs stay left aligned with the first line of each entry.
- Short entries remain on one line when possible.
- Long entries place authors on the second line with dotted leaders before them.
- Page numbers stay at the far right.
- There is no text overlap or missing leader area.

## Common Non-Blocking Warnings
- `fancyhdr \headheight is too small` is usually non-blocking unless the user asks to clean warnings.
- Some `Overfull \hbox` warnings can come from long English terms or formulas in body text; inspect layout before changing unrelated content.
