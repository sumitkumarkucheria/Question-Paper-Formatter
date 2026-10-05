"""
paper_builder.py
-----------------
Builds a Word (.docx) question paper that visually matches the
"MOUNT LITERA ZEE SCHOOL" half-yearly exam template:

  - thin/thick blue page border on every page (color 548DD4)
  - centered bold header block (school name / exam title / subject)
  - two-column Grade|Date and Max Marks|Time row
  - horizontal rule
  - bold "GENERAL INSTRUCTIONS" block, roman-numeral (i)(ii)(iii)...
  - horizontal rule
  - "SECTION A" / "MULTIPLE CHOICE QUESTIONS ( n x m = total Marks )" headings
  - continuously numbered questions (1, 2, 3 ... straight through all sections)
  - MCQ options laid out two-per-line, (A) ... (B) ...
  - code snippets rendered bold in a distinct font (Bahnschrift, like the source)
  - footer: centered page number, fixed school logo bottom-right (static/logo.*)

Also converts the generated .docx to .pdf with LibreOffice (soffice),
so both files can be handed back to the caller.
"""

import copy
import os
import shutil
import subprocess
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BORDER_COLOR = "548DD4"
SCHOOL_NAME = "MOUNT LITERA ZEE SCHOOL, SURAT"   # fixed - always printed in the header
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")


def get_logo_path():
    """The footer logo: static/logo.<png|jpg|jpeg>. Fixed - it is used for every
    paper - until it is replaced from the web page (Exam Details -> School logo)."""
    for ext in ("png", "jpg", "jpeg"):
        p = os.path.join(STATIC_DIR, f"logo.{ext}")
        if os.path.exists(p):
            return p
    return None


LOGO_WIDTH_IN = 1.15
BODY_FONT = "Times New Roman"
CODE_FONT = "Bahnschrift"

SIZE_SCHOOL = 20      # school name
SIZE_TITLE = 14       # exam title / subject / grade-date row
SIZE_SECTION = 17     # "SECTION A"
SIZE_SECTION_SUB = 14 # "MULTIPLE CHOICE QUESTIONS ( ... )"
SIZE_BODY = 14         # question / option text
SIZE_INSTR = 13

# Indentation (inches) - kept tight/shallow so nested content (options,
# sub-parts, wrapped code lines) doesn't creep too far right.
QUESTION_INDENT = 0.24   # wrapped/continuation lines of a question
OPTION_INDENT = 0.26     # MCQ options
SUBPART_INDENT = 0.6     # a) b) c) sub-parts - clearly indented under the question text
INSTR_INDENT = 0.52      # fixed column where every instruction's text starts (fits "(viii)")

# Page geometry. The page border is drawn a fixed distance from the PAGE EDGE
# (BORDER_OFFSET_PT - the outside gap), while the text sits MARGIN_LR from the
# page edge. Gap between border and text = margin - border offset, so making
# the left/right margin smaller tightens ONLY the inside gap.
PAGE_W = 8.27
PAGE_H = 11.69
BORDER_OFFSET_PT = 24    # page edge -> blue border (outside gap, unchanged)
MARGIN_LR = 0.6         # was 0.7 - text edge -> page edge (inches)
RULE_BLEED_PT = MARGIN_LR * 72 - BORDER_OFFSET_PT - 1   # header rule reaches the page border
CHOICE_INDENT = 0.24     # "OR" alternate-question text

ROMAN = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x",
         "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx"]


# --------------------------------------------------------------------------
# low level helpers
# --------------------------------------------------------------------------

def _set_page_border(section):
    sectPr = section._sectPr
    pgBorders = OxmlElement('w:pgBorders')
    pgBorders.set(qn('w:offsetFrom'), 'page')
    for edge, val in (('top', 'thinThickSmallGap'), ('left', 'thinThickSmallGap'),
                       ('bottom', 'thickThinSmallGap'), ('right', 'thickThinSmallGap')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), val)
        el.set(qn('w:sz'), '18')
        el.set(qn('w:space'), str(BORDER_OFFSET_PT))
        el.set(qn('w:color'), BORDER_COLOR)
        pgBorders.append(el)
    sectPr.append(pgBorders)


# elements that must come AFTER <w:pBdr> in a paragraph's properties (schema order)
_AFTER_PBDR = ('w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
               'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
               'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
               'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
               'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
               'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')


def _bottom_rule(paragraph, color=BORDER_COLOR, size=12, space=1):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), str(space))
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    for tag in _AFTER_PBDR:
        nxt = pPr.find(qn(tag))
        if nxt is not None:
            nxt.addprevious(pBdr)
            break
    else:
        pPr.append(pBdr)


def _bleed_to_page_border(paragraph, bleed_pt):
    """Stretch a paragraph (and so its bottom rule) outward past the text
    margins, left and right, so a rule reaches the page border. The first-line
    indent cancels the negative left indent, so the text itself stays put."""
    pf = paragraph.paragraph_format
    pf.left_indent = Pt(-bleed_pt)
    pf.first_line_indent = Pt(bleed_pt)
    pf.right_indent = Pt(-bleed_pt)


def _no_space(paragraph, before=0, after=0, line=1.0):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def _set_run_font(run, font):
    rpr = run._element.get_or_add_rPr()
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rpr.append(rFonts)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)


def _run(paragraph, text, bold=False, size=SIZE_BODY, font=BODY_FONT,
         italic=False, underline=False, color=None):
    r = paragraph.add_run(text)
    r.bold = bold
    r.italic = italic
    r.underline = underline
    r.font.size = Pt(size)
    r.font.name = font
    _set_run_font(r, font)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _para(doc, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    _no_space(p)
    return p


def _add_tab_stop(paragraph, pos_inches):
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(pos_inches), WD_TAB_ALIGNMENT.LEFT)


# --------------------------------------------------------------------------
# header / footer
# --------------------------------------------------------------------------

def _build_header(doc, meta, content_width):
    p = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, SCHOOL_NAME.upper(), bold=True, size=SIZE_SCHOOL)

    p = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, meta["exam_title"].upper(), bold=True, size=SIZE_TITLE)

    p = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
    _run(p, f"SUBJECT: {meta['subject'].upper()}", bold=True, size=SIZE_TITLE)
    p.paragraph_format.space_after = Pt(6)

    # Grade | Date and Max Marks | Time - plain paragraphs with a
    # right-aligned tab stop at the page's content width, so the right
    # side lines up flush with the page border with no table/gridlines.
    def info_line(left_text, right_label, right_value, underline_value=False):
        p = _para(doc)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(content_width), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, left_text, bold=True, size=SIZE_TITLE)
        r = p.add_run(f"\t{right_label}")
        r.bold = True
        r.font.size = Pt(SIZE_TITLE)
        r.font.name = BODY_FONT
        _set_run_font(r, BODY_FONT)
        if right_value:
            rv = p.add_run(right_value)
            rv.bold = True
            rv.underline = underline_value
            rv.font.size = Pt(SIZE_TITLE)
            _set_run_font(rv, BODY_FONT)
        return p

    info_line(f"GRADE: {meta['grade']}", "DATE: ", meta.get("date") or "____/____/______")
    time_val = (meta.get("time") or "").strip()
    marks_line = info_line(f"MAXIMUM MARKS \u2013 {meta['max_marks']}", "TIME: ", time_val,
                           underline_value=bool(time_val))

    # The blue rule is now the bottom border of the Max Marks line itself (no
    # separate empty paragraph), so it hugs the text, and it runs out to the
    # page border on both sides.
    _bottom_rule(marks_line, space=2)
    _bleed_to_page_border(marks_line, RULE_BLEED_PT)


def _build_instructions(doc, instructions):
    p = _para(doc)
    p.paragraph_format.space_before = Pt(6)
    _run(p, "GENERAL INSTRUCTIONS:", bold=True, size=SIZE_TITLE)

    for i, text in enumerate(instructions):
        p = _para(doc)
        p.paragraph_format.left_indent = Inches(INSTR_INDENT)
        p.paragraph_format.first_line_indent = Inches(-INSTR_INDENT)
        # tab stop at the hanging indent: number at the margin, text always
        # starts at the same fixed position, wrapped lines line up under it
        _add_tab_stop(p, INSTR_INDENT)
        _run(p, f"({ROMAN[i] if i < len(ROMAN) else i+1})\t", bold=True, size=SIZE_INSTR)
        _run(p, text, bold=True, size=SIZE_INSTR)

    # Blue rule: a hair-thin paragraph (1pt tall) directly under the last
    # instruction carrying the bottom border, stretched out to the page border
    # on both sides. space_after keeps a gap before the first SECTION heading.
    rule = _para(doc)
    rule.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    rule.paragraph_format.line_spacing = Pt(1)
    rule.paragraph_format.space_after = Pt(6)
    # paragraph-mark font size 1pt so the empty line can't grow taller
    mark_rpr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '2')
    mark_rpr.append(sz)
    rule._p.get_or_add_pPr().append(mark_rpr)
    _bottom_rule(rule, space=0)
    _bleed_to_page_border(rule, RULE_BLEED_PT)


def _add_page_number_footer(doc, content_width):
    """One footer line: page number centered, school logo at the right edge.
    The logo is fixed (static/logo.*) and appears on every page."""
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    _no_space(p)
    tabs = p.paragraph_format.tab_stops
    # the built-in "Footer" style already defines centre/right tab stops at
    # 3.25" and 6.5"; clear them so only our two stops apply
    tabs.add_tab_stop(Inches(3.25), WD_TAB_ALIGNMENT.CLEAR)
    tabs.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.CLEAR)
    tabs.add_tab_stop(Inches(content_width / 2), WD_TAB_ALIGNMENT.CENTER)
    tabs.add_tab_stop(Inches(content_width), WD_TAB_ALIGNMENT.RIGHT)

    p.add_run("\t")
    run = p.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = "PAGE"
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)
    run.font.size = Pt(10)
    run.font.name = BODY_FONT

    logo = get_logo_path()
    if logo:
        p.add_run("\t")
        try:
            p.add_run().add_picture(logo, width=Inches(LOGO_WIDTH_IN))
        except Exception:
            pass


# --------------------------------------------------------------------------
# section / question rendering
# --------------------------------------------------------------------------

def _section_heading(doc, label, title, count, marks_each, total):
    p = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_before = Pt(8)
    _run(p, f"SECTION {label}", bold=True, size=SIZE_SECTION)

    p2 = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
    marks_txt = f"{count} x {marks_each} = {total} Marks" if marks_each == int(marks_each) else \
                f"{count} x {marks_each} = {total} Marks"
    _run(p2, f"{title} ( {marks_txt} )", bold=True, size=SIZE_SECTION_SUB)
    p2.paragraph_format.space_after = Pt(6)


def _question_number_para(doc, number):
    p = _para(doc)
    p.paragraph_format.left_indent = Inches(QUESTION_INDENT)
    p.paragraph_format.first_line_indent = Inches(-QUESTION_INDENT)
    _run(p, f"{number}. ", bold=True, size=SIZE_BODY)
    return p


def _add_text_block(p_first, doc, text, code=False, indent=QUESTION_INDENT):
    """Write (possibly multi-line) text starting in p_first, continuing
    on freshly-created indented paragraphs for subsequent lines."""
    lines = text.split("\n")
    first = True
    p = p_first
    for line in lines:
        if not first:
            p = _para(doc)
            p.paragraph_format.left_indent = Inches(indent)
        _run(p, line, bold=code, size=SIZE_BODY, font=CODE_FONT if code else BODY_FONT)
        first = False
    return p


def _add_options_correct(doc, options, indent=OPTION_INDENT, content_width=PAGE_W - 2 * MARGIN_LR):
    """Render MCQ options two per line: (A) x   (B) y // (C) z  (D) w.
    If a pair's combined text is too long to comfortably fit two columns,
    each option instead gets its own full-width line."""
    letters = ["A", "B", "C", "D", "E", "F"]
    col_gap = 2.9
    # rough capacity check: ~ (available width in inches) * ~9 chars/inch at 14pt
    avail_per_col = max(content_width - indent - col_gap, 1.5)
    max_chars_per_col = int(avail_per_col * 9)

    i = 0
    last_p = None
    while i < len(options):
        a_txt = f"({letters[i]}) {options[i]}"
        b_txt = f"({letters[i+1]}) {options[i+1]}" if i + 1 < len(options) else None
        fits_two_up = b_txt is not None and len(a_txt) <= max_chars_per_col and len(b_txt) <= max_chars_per_col

        if fits_two_up:
            p = _para(doc)
            p.paragraph_format.left_indent = Inches(indent)
            _add_tab_stop(p, indent + col_gap)
            _run(p, a_txt, size=SIZE_BODY)
            _run(p, f"\t{b_txt}", size=SIZE_BODY)
            i += 2
        else:
            p = _para(doc)
            p.paragraph_format.left_indent = Inches(indent)
            _run(p, a_txt, size=SIZE_BODY)
            i += 1
        last_p = p
    return last_p


def _add_sub_parts(doc, sub_parts, indent=SUBPART_INDENT):
    """Sub-parts get their own (a)/(b)/(c)... numbering, indented clearly to
    the right of the question number, with wrapped lines hanging under the
    text (same tab-stop pattern as the general instructions)."""
    letters = "abcdefghijklmnopqrstuvwxyz"
    label_start = QUESTION_INDENT  # "(a)" lines up under the question's own text
    p = None
    for i, sp in enumerate(sub_parts):
        p = _para(doc)
        p.paragraph_format.left_indent = Inches(indent)
        p.paragraph_format.first_line_indent = Inches(label_start - indent)
        _add_tab_stop(p, indent)
        label = letters[i] if i < len(letters) else str(i + 1)
        _run(p, f"({label})\t", size=SIZE_BODY)
        _run(p, sp, size=SIZE_BODY)
    return p


def _add_question_image(doc, image_path, max_width_in=4.3):
    if not image_path or not os.path.exists(image_path):
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _no_space(p, before=4, after=4)
    run = p.add_run()
    try:
        run.add_picture(image_path, width=Inches(max_width_in))
    except Exception:
        pass


def _add_question_table(doc, table_data, indent=QUESTION_INDENT):
    """table_data = {'rows': [[cell, cell, ...], ...], 'has_header': bool}"""
    rows = table_data.get("rows") or []
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        return
    n_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=len(rows), cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # visible thin black grid (this is user question content, not the
    # page-layout table, so real borders make sense here)
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)

    has_header = bool(table_data.get("has_header"))
    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            text = str(row[ci]) if ci < len(row) else ""
            cell = tbl.rows[ri].cells[ci]
            cp = cell.paragraphs[0]
            _no_space(cp)
            _run(cp, text, bold=(has_header and ri == 0), size=SIZE_BODY - 1)

    # small spacer after the table
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _render_question(doc, number, q, content_width=PAGE_W - 2 * MARGIN_LR):
    """
    q: {
      text: str (main prompt, '\n' separated lines),
      code: bool,
      options: [str,...] or [],
      sub_parts: [str,...] or [],
      choice_text: str or None (an "OR" internal-choice alternate question),
      choice_sub_parts: [str,...] or [] (sub-parts of the alternate question),
      image_paths: [str, ...] or [],
      tables: [{"rows": [[...],...], "has_header": bool}, ...] or [],
    }
    """
    p = _question_number_para(doc, number)
    _run(p, q.get("text", "").split("\n")[0], size=SIZE_BODY)
    rest_lines = q.get("text", "").split("\n")[1:]
    last_p = p
    for line in rest_lines:
        lp = _para(doc)
        lp.paragraph_format.left_indent = Inches(QUESTION_INDENT)
        _run(lp, line, bold=q.get("code", False), size=SIZE_BODY,
             font=CODE_FONT if q.get("code") else BODY_FONT)
        last_p = lp

    for img_path in (q.get("image_paths") or []):
        _add_question_image(doc, img_path)

    for tbl_data in (q.get("tables") or []):
        _add_question_table(doc, tbl_data)

    if q.get("options"):
        last_p = _add_options_correct(doc, q["options"], content_width=content_width) or last_p

    if q.get("sub_parts"):
        last_p = _add_sub_parts(doc, q["sub_parts"]) or last_p

    if q.get("choice_text"):
        orp = _para(doc, WD_ALIGN_PARAGRAPH.CENTER)
        orp.paragraph_format.space_before = Pt(4)
        orp.paragraph_format.space_after = Pt(4)
        _run(orp, "OR", bold=True, italic=True, size=SIZE_BODY)
        cp = _para(doc)
        cp.paragraph_format.left_indent = Inches(CHOICE_INDENT)
        _run(cp, q["choice_text"], size=SIZE_BODY)
        last_p = cp
        if q.get("choice_sub_parts"):
            last_p = _add_sub_parts(doc, q["choice_sub_parts"])

    # No extra gap before the next question - the question number itself
    # (bold "1.", "2." ...) already reads as a clear break, so removing this
    # saves paper over many questions.
    last_p.paragraph_format.space_after = Pt(0)


# --------------------------------------------------------------------------
# public entry point
# --------------------------------------------------------------------------

def build_docx(paper, out_path):
    """
    paper = {
      "exam_title": str,
      "subject": str,
      "grade": str,
      "max_marks": int/str,
      "date": str,
      "time": str,
      "instructions": [str, ...],
      "sections": [
        {
          "label": "A",
          "title": "MULTIPLE CHOICE QUESTIONS",
          "marks_each": 1,
          "questions": [ {text, code, options, sub_parts, choice_text}, ... ]
        }, ...
      ]
    }
    """
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(PAGE_W)
    section.page_height = Inches(PAGE_H)  # A4, matches source
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(MARGIN_LR)
    section.right_margin = Inches(MARGIN_LR)
    _set_page_border(section)

    doc.styles['Normal'].font.name = BODY_FONT
    doc.styles['Normal'].font.size = Pt(SIZE_BODY)

    content_width = PAGE_W - 2 * MARGIN_LR  # page width - left margin - right margin

    _add_page_number_footer(doc, content_width)

    meta = {
        "exam_title": paper["exam_title"],
        "subject": paper["subject"],
        "grade": paper["grade"],
        "max_marks": paper["max_marks"],
        "date": paper.get("date", ""),
        "time": paper.get("time", ""),
    }
    _build_header(doc, meta, content_width)
    _build_instructions(doc, paper["instructions"])

    qnum = 1
    for sec in paper["sections"]:
        count = len(sec["questions"])
        marks_each = sec["marks_each"]
        total = round(count * marks_each, 2)
        if total == int(total):
            total = int(total)
        _section_heading(doc, sec["label"], sec["title"], count, marks_each, total)
        for q in sec["questions"]:
            _render_question(doc, qnum, q, content_width=content_width)
            qnum += 1

    doc.save(out_path)
    return out_path


def _convert_via_word_windows(docx_path, pdf_path, attempts=3):
    """Converts via a dedicated, independent Word automation instance.

    Using win32com's DispatchEx (rather than docx2pdf's plain Dispatch)
    starts a fresh, invisible Word process for this conversion instead of
    attaching to whatever Word window the user already has open (e.g. from
    opening a previously downloaded paper to check it). Attaching to an
    existing window is the likely cause of "regenerate right after opening
    the download" failing: a document open in that window, or a dialog on
    it, blocks the automation call. A retry loop also rides out Word's
    transient "call was rejected because the callee is busy" (RPC_E_CALL_
    REJECTED) error, which shows up when Word is asked to do something
    again very soon after the previous call.
    """
    import time
    import pythoncom
    import win32com.client

    last_err = None
    for attempt in range(attempts):
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(
                os.path.abspath(docx_path), ReadOnly=True, AddToRecentFiles=False
            )
            try:
                doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
            finally:
                doc.Close(False)
            return
        except Exception as e:
            last_err = e
            time.sleep(1.5 * (attempt + 1))  # give a busy Word a moment, then retry
        finally:
            if word is not None:
                try:
                    word.Quit(SaveChanges=0)
                except Exception:
                    pass
            pythoncom.CoUninitialize()
    raise last_err


def convert_to_pdf(docx_path, out_dir):
    """
    Converts docx -> pdf. Tries, in order:
      1. A dedicated MS Word automation instance (Windows, if Word + pywin32
         are installed) - independent of any Word window the user has open.
      2. docx2pdf (Windows/Mac, if Word is installed and pywin32 wasn't
         available directly) as a secondary path.
      3. LibreOffice (soffice) headless, checking a few common install paths
         in addition to PATH.
    Raises RuntimeError with guidance if none are available, including the
    real underlying error from whichever Word attempt was actually tried
    (previously this was silently discarded, hiding the real cause).
    """
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    last_word_error = None

    # --- Option 1: dedicated Word automation instance (Windows only) ---
    if os.name == "nt":
        try:
            _convert_via_word_windows(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                return pdf_path
        except ImportError:
            pass  # pywin32 not installed - fall through
        except Exception as e:
            last_word_error = e

    # --- Option 2: MS Word via docx2pdf (uses your existing Word install) ---
    if not os.path.exists(pdf_path):
        try:
            from docx2pdf import convert as _word_convert
            _word_convert(docx_path, pdf_path)
            if os.path.exists(pdf_path):
                return pdf_path
        except ImportError:
            pass  # docx2pdf not installed - fall through to LibreOffice
        except Exception as e:
            last_word_error = e

    # --- Option 2: LibreOffice headless ---
    # Each call gets its own throwaway user-profile directory. Without this,
    # every conversion shares the one default LibreOffice profile, which the
    # first run locks; a second run (e.g. after a page refresh) then fails
    # with a profile-in-use error, especially if the first soffice process
    # hadn't fully exited yet. An isolated profile per call avoids that.
    import tempfile
    candidates = ["soffice", "libreoffice"]
    if os.name == "nt":
        candidates += [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    last_error = None
    for cmd in candidates:
        profile_dir = tempfile.mkdtemp(prefix="soffice_profile_")
        try:
            result = subprocess.run(
                [cmd, "--headless", "--norestore",
                 f"-env:UserInstallation=file://{profile_dir}",
                 "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                capture_output=True, text=True, timeout=120
            )
            if os.path.exists(pdf_path):
                return pdf_path
            last_error = result.stderr or result.stdout
        except (FileNotFoundError, OSError) as e:
            last_error = str(e)
            continue
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)

    raise RuntimeError(
        "Could not convert to PDF. Neither MS Word nor LibreOffice (soffice) "
        "could be used.\n"
        "- To use your existing MS Word: pip install docx2pdf pywin32\n"
        "- Or install LibreOffice: https://www.libreoffice.org/download/download/\n"
        f"Word error: {last_word_error}\n"
        f"LibreOffice error: {last_error}"
    )


def default_instructions(sections, extra=None):
    """Auto-generate the GENERAL INSTRUCTIONS block from the section layout,
    mirroring the phrasing/order used in the source template."""
    total_q = sum(len(s["questions"]) for s in sections)
    n_sections = len(sections)
    letters = [s["label"] for s in sections]
    lines = []
    lines.append(f"This question paper contains {total_q} questions.")
    lines.append("All questions are compulsory. However, internal choices have been "
                  "provided in some questions. Attempt only one of the choices in such questions.")
    lines.append(f"The paper is divided into {n_sections} Sections \u2013 "
                 + ", ".join(letters[:-1]) + (" and " + letters[-1] if len(letters) > 1 else letters[0]) + ".")

    qnum = 1
    for s in sections:
        c = len(s["questions"])
        start, end = qnum, qnum + c - 1
        rng = f"{start}" if start == end else f"{start} to {end}" if c > 2 else f"{start} and {end}"
        me = s["marks_each"]
        me_txt = int(me) if me == int(me) else me
        lines.append(f"Section {s['label']}, consists of {c} question{'s' if c != 1 else ''} "
                     f"({rng}). Each question carries {me_txt} mark{'s' if me_txt != 1 else ''}.")
        qnum += c
    lines.append("In case of MCQs, text of the correct answer should also be written.")
    if extra:
        lines.extend(extra)
    return lines
