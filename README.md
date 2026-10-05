# Question Paper Generator (qpgen)

A self-hosted Flask web app that builds school exam question papers as **Word (`.docx`) and PDF** files in a fixed, professional template: blue page border, centered header block, numbered general instructions, lettered sections with marking formulas, continuously numbered questions, and a footer with page number and school logo.

Teachers fill in a 5-step wizard in the browser. The server assembles the document with `python-docx` and converts it to PDF. No database, no external services, no frontend build step.

---

## Features

- **5-step wizard UI** (vanilla JS, single HTML file): Exam Details, Instructions, Marking Scheme, Questions, Review & Generate
- **Template-accurate output**
  - Blue page border (`#548DD4`) on every page
  - Centered header: school name, exam title, subject
  - Grade / Date and Max Marks / Time row, with horizontal rules
  - Bold `GENERAL INSTRUCTIONS` block numbered (i), (ii), (iii)...
  - `SECTION A / B / C...` headings with the formula `( n x m = total Marks )`
  - Questions numbered continuously across sections
  - Footer with centered page number and school logo (bottom-right)
- **Question types**: MCQ, short answer, long answer, case study
- **MCQ layout**: options in two columns, auto-wrapping to their own line when too long
- **Per-question extras**: code-block mode (bold monospace), sub-parts (a, b, c...), internal-choice "OR" alternate, inline images, simple data tables
- **Auto-generated instructions** derived from your marking scheme (question counts, per-section ranges, marks per question), editable afterwards
- **Live marks check**: banner compares planned total against Maximum Marks, with a one-click balance button
- **Logo management**: upload a PNG/JPG logo, or reset to the bundled default
- **Multi-user on LAN**: served with Waitress (8 threads); PDF conversion is serialized with a lock

---

## Tech Stack

| Layer | Tool |
|---|---|
| Backend | Python 3.9+, Flask |
| Document engine | python-docx (with raw OOXML for borders, rules, tab stops) |
| PDF conversion | MS Word automation (Windows), `docx2pdf`, or LibreOffice headless, tried in that order |
| Server | Waitress (falls back to Flask dev server if not installed) |
| Frontend | Vanilla HTML/CSS/JS, no build step |

---

## Project Structure

```
qpgen/
├── app.py              # Flask routes: UI, logo upload, asset upload, auto-instructions, generate, download
├── paper_builder.py    # DOCX engine (layout, fonts, border, tables, images) + PDF conversion
├── templates/
│   └── index.html      # Single-page wizard UI
├── static/
│   ├── logo.png            # Active footer logo
│   ├── default_logo.png    # Bundled default, restored by "reset logo"
│   └── uploads/            # Question images uploaded from the wizard
├── generated/          # Output .docx and .pdf files
├── requirements.txt
└── README.md
```

---

## Requirements

- Python 3.9 or newer
- **One** of the following for PDF export:
  - **LibreOffice** with `soffice` on PATH (all platforms; recommended for Linux/macOS)
  - **Microsoft Word** + `pywin32` (Windows) or `docx2pdf` (Windows/Mac)

If no converter is available, the `.docx` is still built but the request returns an error because PDF conversion is part of the same call (see [Known Limitations](#known-limitations)).

Install LibreOffice:

```bash
# Debian / Ubuntu
sudo apt install libreoffice

# macOS (Homebrew)
brew install --cask libreoffice

# Windows: download installer from https://www.libreoffice.org/
```

---

## Installation

```bash
git clone <your-repo-url>
cd qpgen

python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate

pip install -r requirements.txt
```

## Run

```bash
python app.py
```

Open **http://localhost:5000**. The server binds to `0.0.0.0`, so other machines on the same network can reach it at `http://<host-ip>:5000`.

Change the port:

```bash
PORT=8080 python app.py           # Windows (PowerShell): $env:PORT=8080; python app.py
```

---

## Usage

1. **Exam Details**: exam type, subject, grade, maximum marks, date, time.
2. **Instructions**: auto-generate from the marking scheme, or write your own line by line.
3. **Marking Scheme**: add sections with a title, question type, number of questions, and marks per question.
4. **Questions**: enter question text, MCQ options, sub-parts, optional "OR" alternates, images, and tables. Toggle *code block* for Python snippets.
5. **Review & Generate**: verify the summary, click Generate, download the `.docx` and `.pdf`.

Output files are written to `generated/` as `<subject>_<exam-title>_<token>.docx|pdf`.

---

## API Endpoints

| Method | Route | Purpose |
|---|---|---|
| GET | `/` | Wizard UI |
| GET | `/logo` | Serves the current footer logo (no-store cache) |
| POST | `/api/upload_logo` | Replace the footer logo (PNG/JPG) |
| POST | `/api/reset_logo` | Restore the bundled default logo |
| POST | `/api/upload_asset` | Upload an image for use inside a question (PNG/JPG/GIF/BMP) |
| POST | `/api/auto_instructions` | Build the general-instructions list from sections |
| POST | `/api/generate` | Validate payload, build DOCX, convert to PDF, return download URLs |
| GET | `/download/<file>` | Download a generated file |

### `/api/generate` payload (shape)

```json
{
  "exam_title": "Half Yearly Examination",
  "subject": "Computer Science",
  "grade": "XII",
  "max_marks": 70,
  "date": "10-10-2026",
  "time": "3 Hours",
  "instructions": ["..."],
  "sections": [
    {
      "label": "A",
      "title": "MULTIPLE CHOICE QUESTIONS",
      "qtype": "mcq",
      "marks_each": 1,
      "questions": [
        {
          "text": "Which keyword defines a function in Python?",
          "code": false,
          "options": ["def", "func", "define", "lambda"],
          "sub_parts": [],
          "use_choice": false,
          "images": [],
          "tables": []
        }
      ]
    }
  ]
}
```

Required fields: `exam_title`, `subject`, `grade`, `max_marks`, and at least one section with at least one non-empty question.

---

## Customization

Layout constants live at the top of `paper_builder.py`:

| Constant | Default | Controls |
|---|---|---|
| `SCHOOL_NAME` | `MOUNT LITERA ZEE SCHOOL, SURAT` | Header school name (hardcoded) |
| `BORDER_COLOR` | `548DD4` | Page border and rule color |
| `BODY_FONT` / `CODE_FONT` | Times New Roman / Bahnschrift | Body and code fonts |
| `SIZE_*` | 13 to 20 pt | Per-element font sizes |
| `MARGIN_LR`, `BORDER_OFFSET_PT` | 0.5 in, 24 pt | Page geometry |
| `LOGO_WIDTH_IN` | 1.15 | Footer logo width |

Page size is A4 (8.27 x 11.69 in).

To adapt this for another school: change `SCHOOL_NAME`, replace the logo from the UI, and adjust colors and fonts.

---

## Known Limitations

Read these before deploying beyond a trusted LAN.

- **No authentication.** Anyone who can reach the port can generate papers, upload files, and replace the logo.
- **Logo is global.** Uploading a logo changes it for every user and every future paper, not per user or per session.
- **School name is hardcoded** in `paper_builder.py`; it cannot be changed from the UI.
- **Default instructions are CS-specific.** The auto-generated block always includes "All programming questions are to be answered using Python Language only." Edit it out for other subjects (`default_instructions()` in `paper_builder.py`).
- **PDF failure fails the whole request.** `build_docx` and `convert_to_pdf` run in one try block, so a missing converter returns a 500 error even though the `.docx` was written to `generated/`.
- **Max Marks is not enforced server-side.** The UI shows a mismatch banner, but `/api/generate` does not reject a paper whose section totals differ from `max_marks`.
- **No cleanup.** Files in `generated/` and `static/uploads/` accumulate indefinitely.
- **Font fidelity.** `Bahnschrift` is a Windows font. On Linux/macOS, LibreOffice substitutes another font, so code snippets in the PDF will look different from Word on Windows.
- **Single-conversion bottleneck.** PDF conversion is serialized by a lock. Fine for 5 to 10 occasional users, not for heavy concurrent use.
- **No saved drafts.** Wizard state lives in browser memory; refreshing the page loses it.
- **Do not expose to the public internet** as-is: no rate limiting, no CSRF protection, no input quotas beyond the 16 MB upload cap.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Could not convert to PDF" | Install LibreOffice and confirm `soffice --version` works in the same shell, or install MS Word + `pywin32` |
| PDF conversion hangs on Linux | Run `soffice --headless --version` once manually; check for stuck `soffice` processes |
| Not reachable from other PCs | Allow the port through the host firewall; use the host's LAN IP, not `localhost` |
| Logo not updating in the preview | Hard refresh; the `/logo` route sends `no-store`, so also check the uploaded file is PNG/JPG |
| `waitress not installed` warning | `pip install waitress`; the Flask dev server fallback is not suitable for multiple users |

---

## Roadmap 

- Per-user auth and per-user logo / school profile
- Configurable school name and template from the UI
- Save/load paper drafts (JSON) and a question bank
- Answer key and marking scheme export
- Server-side marks validation
- Scheduled cleanup of `generated/` and `uploads/`
- Docker image bundling LibreOffice and fonts

---

## License

No license.
