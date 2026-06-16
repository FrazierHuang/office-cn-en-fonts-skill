---
name: office-cn-en-fonts-skill
description: Enforce Office output typography and spacing for Chinese/English mixed documents. Use when Codex creates, edits, formats, or post-processes Word (.docx), Excel (.xlsx), or PowerPoint (.pptx) files and the output must use SimSun/宋体 for Chinese text and Times New Roman for English text, with no spaces between adjacent Chinese and English text; for Word and Excel outputs, all font color must be black and tables must have no fill/shading color.
---

# Office Chinese/English Fonts Skill

## Core Rules

Apply these rules to every generated Office artifact unless the user explicitly overrides them:

- Chinese text: use `宋体` / `SimSun`.
- English letters, numbers, and Western punctuation: use `Times New Roman`.
- Word and Excel: set all visible text to black.
- Word and Excel: remove table or cell fill colors; tables must use no shading/fill.
- PowerPoint: enforce the Chinese/English font pairing, but do not force black text unless the user asks.
- Remove spaces between adjacent Chinese and English/numeric text, e.g. write `河蚬As` and `As暴露`, not `河蚬 As` or `As 暴露`.
- For Word and PowerPoint, remove those boundary spaces even when the phrase is split across multiple runs/text nodes.

## Creation Workflow

When creating new files, set formatting during generation instead of relying only on cleanup:

1. For DOCX, set run fonts with OOXML `w:rFonts`:
   - `w:eastAsia="宋体"` or `SimSun`
   - `w:ascii="Times New Roman"`
   - `w:hAnsi="Times New Roman"`
   - `w:cs="Times New Roman"`
   Also set `w:color w:val="000000"` for Word outputs.
2. For XLSX, prefer rich text runs when a cell mixes Chinese and English. Assign Chinese runs to `宋体` and Western runs to `Times New Roman`. Set font color to black and avoid fill styles.
3. For PPTX, set text run properties with Latin font `Times New Roman` and East Asian font `宋体` / `SimSun`.
4. When writing mixed Chinese-English text, do not insert a space at the Chinese/English boundary.
5. Do not use colored header rows, banded rows, or shaded table cells in Word or Excel unless the user explicitly asks for them.

## Post-Processing

After creating or editing an Office file, run the bundled cleanup script when possible:

```bash
python3 /path/to/office-cn-en-fonts-skill/scripts/enforce_office_fonts.py --in-place file.docx file.xlsx file.pptx
```

Use `--in-place` for final artifacts. Without `--in-place`, the script writes a sibling file with `_fontfixed` before the extension. Use `--backup` with `--in-place` when preserving originals matters. Use `--check` to audit files without modifying them. Use `--recursive` when the input is a folder.

The script:

- updates DOCX run/style fonts and forces black Word text;
- removes DOCX table-cell shading;
- removes spaces between adjacent Chinese and English/numeric text inside and across OOXML text nodes;
- updates XLSX font colors, resets fills, and rewrites shared/inline strings into rich text runs split into Chinese and non-Chinese segments;
- updates PPTX run/theme font declarations for Latin and East Asian text.
- returns a non-zero exit code in `--check` mode if any file still has formatting issues.

Examples:

```bash
python3 /path/to/office-cn-en-fonts-skill/scripts/enforce_office_fonts.py --check outputs
python3 /path/to/office-cn-en-fonts-skill/scripts/enforce_office_fonts.py --recursive --in-place --backup outputs
```

## Verification

Before saying the artifact is finished:

1. Inspect or unzip the generated Office file when feasible.
2. Confirm DOCX contains `Times New Roman` and either `宋体` or `SimSun` in `word/*.xml`.
3. Confirm XLSX `xl/styles.xml` has black font color and no custom fills, and mixed strings are represented with rich text runs when applicable.
4. Confirm PPTX slide/theme XML contains Latin `Times New Roman` and East Asian `宋体` or `SimSun`.
5. Search representative text to confirm Chinese-English boundaries do not contain spaces.
6. If the file was rendered to PDF or preview images, visually check that tables are unfilled and Word/Excel text is black.
