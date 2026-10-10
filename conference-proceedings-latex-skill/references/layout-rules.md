# Layout Rules

## Data Source And Generated Files
- The Excel source is typically `摘要信息.xls`.
- `generate_tex.js` reads Excel metadata and writes `toc_content.tex` and `papers_content.tex`.
- `\tocentry{paperId}{labelId}{title}{authors}` is the normal TOC call; use an additional macro such as `\tocentrytwo` only when the generator intentionally emits long entries differently.

## Body Rules
- Under author affiliations, show corresponding-author email as `\textsuperscript{*}email`.
- Do not show the literal prefix `通讯作者邮箱：` before the email.
- Keep author names, affiliation numbers, and corresponding-author star consistent with the existing `\printpaper` API.
- Avoid introducing stray boolean/text values such as `false` into generated body output.

## Escaping And Content Sanitation
- Escape bare `&` in generated LaTeX unless it is intentionally part of a LaTeX alignment context.
- Convert common HTML entities before LaTeX escaping when appropriate.
- Preserve existing safeguards for `%`, `_`, `#`, `~`, and unsupported characters.

## Table Of Contents Rules
- Keep paper ID visible at the left.
- Short entries should fit on one line: `paperId + title + dotted leader + authors + page`.
- Long entries should use two lines: first line has `paperId + title + dotted leader`; second line has dotted leader before the author names, then `authors（page）` at the right.
- Page number must stay at the far right. Do not let page numbers fall to the left or become separated from the right-side result.
- When authors move to the second line, the author names should end next to the page number; the space before the names should be filled by dotted leaders.
- Avoid layouts where title text overlaps authors, author names overlap page numbers, or paper IDs align with the second line instead of the first.

## Visual Preference
- Use Chinese full-width parentheses for page references: `（23）`.
- Dotted leaders should use the existing `\leaders\hbox to .5em{\hss.\hss}\hfill` style unless the project already uses another leader style.
- Keep TOC density readable; do not add large vertical gaps unless needed to prevent overlap.
