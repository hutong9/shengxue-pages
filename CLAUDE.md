This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Project Overview

This is a LaTeX-based academic conference proceedings generator for the 2026 National Acoustics Conference (2026年全国声学大会). The system processes Excel data containing paper abstracts and generates a formatted PDF proceedings book with table of contents, hyperlinks, and proper academic formatting.

## Architecture

The project consists of three main components:

1. **Data Processing (`generate_tex.js`)**: Node.js script that reads Excel data and generates LaTeX content files
2. **LaTeX Templates**: Two main templates for different output formats
   - `main.tex`: Single paper example template
   - `main_full.tex`: Complete proceedings with generated content
3. **Generated Content Files**:
   - `papers_content.tex`: All paper entries in LaTeX format
   - `toc_content.tex`: Table of contents with hyperlinks

## Build Process

The complete build workflow requires these steps in order:

1. **Generate LaTeX content from Excel data**:
   ```bash
   node generate_tex.js
   ```
   This reads `摘要信息.xls` and generates `papers_content.tex` and `toc_content.tex`

2. **Compile the complete proceedings** (requires two passes for proper cross-references):
   ```bash
   pdflatex main_full
   pdflatex main_full
   ```
   Output: `main_full.pdf`

3. **For single paper testing**:
   ```bash
   pdflatex main
   pdflatex main
   ```
   Output: `main.pdf`

## Key Dependencies

- Node.js with `xlsx` package for Excel processing
- pdfLaTeX with CJK support for Chinese text
- Required LaTeX packages: `ctexart`, `newtxtext`, `newtxmath`, `hyperref`, `fancyhdr`, `geometry`

## Data Structure

The Excel file (`摘要信息.xls`) contains columns for:
- Paper metadata: `论文专题`, `论文编号`, `论文标题`
- Authors: `作者姓名1-12`, `作者单位1-12`, `通讯作者姓名`
- Content: `摘要文本`, `论文关键字`

## LaTeX Formatting Features

- Custom font size commands (`\wuhao`, `\xiaosi`, etc.) for Chinese typography
- Author affiliation numbering with automatic deduplication
- Corresponding author marking with `\ca` (asterisk superscript)
- HTML entity conversion to LaTeX commands
- Special character escaping for LaTeX compatibility
- Hyperlinked table of contents with dotted leaders

## Common Issues

The table of contents generation has formatting issues where single-line titles don't get proper dotted leaders to align page numbers to the right margin. This affects entries like "G0005" and other papers with titles that fit exactly on one line.