# Question Paper Generator
It does not generate questions but format them in clean way!
A standalone Python/Flask web app that builds a school question paper
(Word `.docx` + `.pdf`) in the same visual style as a bordered, sectioned
exam template: blue page border, centered header block (school name /
exam title / subject), Grade–Date and Max Marks–Time row, a horizontal
rule, a bold "GENERAL INSTRUCTIONS" block numbered (i)(ii)(iii)…,
"SECTION A / B / C…" headings with the marking formula
`( n x m = total Marks )`, continuously numbered questions, two-column
MCQ options, bold code snippets, and a footer with page number + your
school logo.

## 1. Requirements

- Python 3.9+
- [LibreOffice](https://www.libreoffice.org/) installed and on your PATH
  as `soffice` (used only to convert the generated `.docx` to `.pdf`).
  - Windows/Mac: install LibreOffice normally.
  - Debian/Ubuntu: `sudo apt install libreoffice`
  - macOS (Homebrew): `brew install --cask libreoffice`

## 2. Setup

```bash
cd qpgen
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 3. Run

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

## 4. Using the wizard

1. **Exam Details** — type of examination, subject, grade, maximum
   marks, date, time, and an optional school logo (shown bottom-right
   of every page, like the source template).
2. **Instructions** — auto-generate the "GENERAL INSTRUCTIONS" block
   from your marking scheme (mirrors the source paper's phrasing), or
   write your own line by line.
3. **Marking Scheme** — add sections (A, B, C…), each with a title
   (e.g. "MULTIPLE CHOICE QUESTIONS"), a question type (MCQ / short /
   long / case-study), how many questions, and marks per question. A
   live banner shows whether the planned total matches your Maximum
   Marks, with a one-click button to balance it.
4. **Questions** — enter each question's text (toggle "code block" to
   render it bold/monospace like the source paper's Python snippets),
   MCQ options (auto-wraps long options to their own line if two won't
   fit side by side), optional sub-parts (a, b, c…) for structured
   questions, an optional internal "OR" alternate question, and you can
   insert an image (diagram/screenshot) and/or a simple data table
   directly into any question.
5. **Review & Generate** — check the summary, then click **Generate**.
   You'll get download links for both the `.docx` and the `.pdf`.

Generated files are written to `qpgen/generated/`; uploaded logos to
`qpgen/static/uploads/`.

## 5. Files

- `app.py` — Flask routes (UI, logo upload, auto-instructions, generate, download).
- `paper_builder.py` — the docx-building engine (page border, fonts,
  sizes, layout) and the LibreOffice PDF conversion call. This is
  where to tweak fonts/sizes/colors if you want to match a different
  template exactly.
- `templates/index.html` — the single-page wizard UI (vanilla JS, no
  build step, no external services).

## 6. Notes

- If LibreOffice isn't installed, the Word file will still generate —
  only the PDF conversion will fail. Install LibreOffice (or open the
  `.docx` and use "Save as PDF" from Word yourself) to get the PDF.
- The layout constants (font sizes, border color `548DD4`, margins)
  live at the top of `paper_builder.py` — edit them there to match a
  different school's template.
