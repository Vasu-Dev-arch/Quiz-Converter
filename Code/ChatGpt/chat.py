#!/usr/bin/env python3
"""
QuizFormatter CLI (no GUI)

Usage:
    python quizformatter_cli.py input.docx output.docx

What it does:
- Parses a .docx containing mixed-format questions (inline options, option-lines, Answer:, Explanation:, Tamil/unicode etc.)
- Produces output .docx where each question becomes a single table with exactly 8 rows and 3 logical columns,
  with row-wise merges matching the professor-required layout.

Notes:
- Requires: python-docx
- Install: pip install python-docx
"""

import sys
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import unicodedata
from typing import List, Optional
from docx import Document

# --------------------------
# Helpers
# --------------------------
def normalize_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    # Normalize Unicode (preserve Tamil and special characters)
    s = unicodedata.normalize("NFC", s)
    # Remove zero-width spaces etc.
    s = s.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
    # Normalize whitespace
    s = re.sub(r'\r\n|\r', '\n', s)
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\n\s+\n', '\n\n', s)
    return s.strip()

SEPARATOR_RE = re.compile(r'^[\-\—\–\*\_]{1,}\s*$')  # lines that are separators like --- — ——, ***, ___

HEADING_PATTERNS = [
    r'^unique questions',
    r'^unique questions with answers',
    r'^unique questions with answers and explanations',
    r'^paper\s*\d+',
    r'^paper\s*\d+\s+unique questions',
    r'^selected questions',
    r'^selected questions on',
    r'^questions? on',
    r'^index\b',
    r'^contents\b',
]

def is_heading_line(line: str) -> bool:
    if not line:
        return False
    ln = line.strip().lower()
    # very short lines that are likely headings (but not remove short questions ending with '?')
    for pat in HEADING_PATTERNS:
        if re.match(pat, ln):
            return True
    # If it contains 'unique questions' or 'paper' or 'answers and explanations' treat as heading
    if 'unique questions' in ln or 'answers and explanations' in ln:
        return True
    # lines like "Paper 1" or "Paper 1 Unique Questions" are headings
    if re.match(r'^paper\s*\d+\b', ln):
        return True
    return False

# --------------------------
# Read & split into blocks
# --------------------------
def read_docx_paragraphs(path: str) -> List[str]:
    doc = Document(path)
    paras = [normalize_text(p.text) for p in doc.paragraphs]
    return paras

def group_paragraphs_into_blocks(paragraphs: List[str]) -> List[List[str]]:
    """
    Groups consecutive non-separator paragraphs into blocks.
    Separators: blank lines or lines matching SEPARATOR_RE.
    Returns list of blocks; each block is a list of paragraph strings.
    """
    blocks = []
    current = []
    for p in paragraphs:
        if p is None:
            continue
        stripped = p.strip()
        if stripped == "" or SEPARATOR_RE.match(stripped):
            if current:
                blocks.append(current)
                current = []
            continue
        # treat some single dash used as separator
        if stripped in ('—', '--', '---', '–'):
            if current:
                blocks.append(current)
                current = []
            continue
        current.append(p)
    if current:
        blocks.append(current)
    return blocks

# --------------------------
# Parsing a block
# --------------------------
LABEL_OPTIONS_RE = re.compile(r'Options?\s*[:\-]?\s*', re.I)
INLINE_LABEL_RE = re.compile(r'[\(\[]?([A-Da-d])[\)\].]?', re.I)
OPTION_LINE_RE = re.compile(r'^\s*[\(\[]?([A-Da-d])[\)\].]?\s*(.+)', re.I)
ANSWER_LINE_RE = re.compile(r'\b(Answer|Ans|Correct|Key|Correct option)\b', re.I)
EXPL_LABEL_RE = re.compile(r'\b(Explanation|Solution|Explanatory)\b', re.I)

def remove_leading_headings(paras: List[str]) -> List[str]:
    # Remove initial heading paragraphs like "Unique Questions..." or "Paper 1..."
    while paras and is_heading_line(paras[0]):
        paras.pop(0)
    return paras

def extract_inline_options_from_text(s: str) -> Optional[List[str]]:
    """
    If the text contains inline options like '(a) ... (b) ... (c) ... (d) ...'
    returns list of option texts (in order a..d). Otherwise returns None.
    """
    # find label occurrences
    matches = list(re.finditer(r'[\(\[]?([A-Da-d])[\)\].]?', s))
    if not matches or len(matches) < 2:
        return None
    opts = []
    for i, m in enumerate(matches):
        start = m.end()
        end = matches[i+1].start() if i+1 < len(matches) else len(s)
        opt = s[start:end].strip()
        # remove trailing punctuation introduced by sentence punctuation
        opt = opt.strip(' ,;:.-')
        opts.append(opt)
    # ensure at most 4
    return [o for o in opts][:4]

def parse_block(paras: List[str], debug_index: int = 0) -> Optional[dict]:
    """
    Parse a block (list of paragraphs) into question dict or return None if not a question.
    returns:
        {
            'question': str,
            'options': [oA,oB,oC,oD],  # always 4 strings
            'answer': 'a'|'b'|'c'|'d',
            'assumed': bool,
            'explanation': str,
            'raw_block': str
        }
    """
    paras = [p for p in paras if p.strip() != ""]
    if not paras:
        return None
    paras = remove_leading_headings(paras)
    if not paras:
        return None
    block_text = "\n".join(paras).strip()

    # Heuristic: if block doesn't contain 'Options' or option markers or 'Answer' and is extremely short, skip
    if (not re.search(r'Options?', block_text, re.I)
        and not re.search(r'[\(\[]?[A-Da-d][\)\].]?', block_text)
        and not re.search(r'\b(Answer|Ans|Correct)\b', block_text, re.I)
        and len(block_text) < 30
        and '?' not in block_text):
        # likely heading or noise
        return None

    # 1) Extract explanation (if any) at the end using label
    explanation = ""
    expl_search = re.search(r'(Explanation|Solution|Explanatory)\s*[:\-]?\s*(.*)$', block_text, re.I | re.S)
    if expl_search:
        explanation = expl_search.group(2).strip()
        # remove explanation from block_text
        block_text = block_text[:expl_search.start()].strip()

    # 2) Extract Answer line (if present)
    raw_answer_text = None
    answer_letter = None
    ans_match = re.search(r'\b(Answer|Ans|Correct|Key|Correct option)\s*[:\-]?\s*([^\n\r]*)', block_text, re.I)
    if ans_match:
        raw_answer_text = ans_match.group(2).strip()
        # remove answer portion from block_text
        block_text = block_text[:ans_match.start()].strip()
        # try letter
        m = re.search(r'[\(\[]?([A-Da-d])[\)\].]?', raw_answer_text)
        if m:
            answer_letter = m.group(1).lower()
        else:
            # maybe it's the option text; we'll resolve later by fuzzy matching
            answer_letter = None

    # 3) Extract options:
    options = []

    # 3a) If there's an explicit "Options:" chunk, handle inline options after that token
    opt_token_match = re.search(r'Options?\s*[:\-]?\s*(.*)$', block_text, re.I | re.S)
    if opt_token_match:
        options_area = opt_token_match.group(1).strip()
        # try inline extraction
        inline_opts = extract_inline_options_from_text(options_area)
        if inline_opts:
            options = inline_opts
            # remove the Options: ... portion from question_text
            block_text = block_text[:opt_token_match.start()].strip()
        else:
            # maybe the paragraph with 'Options:' had content that is same paragraph combined with other text,
            # fall back to scanning entire paragraph lines for options
            # we'll continue to look for labeled option lines below
            pass

    # 3b) If no options found yet, search for lines that start with 'a.' '(a)' etc.
    if not options:
        for p in paras:
            m = OPTION_LINE_RE.match(p)
            if m:
                options.append(m.group(2).strip())

    # 3c) If still no options, try to find inline labeled options across the whole block_text
    if not options:
        inline_opts_whole = extract_inline_options_from_text(block_text)
        if inline_opts_whole:
            options = inline_opts_whole
            # remove them from question area by removing everything from first label onward
            # find first label in block_text
            mfirst = re.search(r'[\(\[]?[A-Da-d][\)\].]?', block_text)
            if mfirst:
                block_text = block_text[:mfirst.start()].strip()

    # 3d) If options found but are fewer than 4, they will be padded later.
    # 3e) If options still empty, as a last resort, attempt to token-split a single-line using 'a)' or 'a.' markers
    if not options:
        # try splitting on ' a) ' ' b) ' pattern
        msplit = re.split(r'\s+[A-Da-d][\)\.\]]\s+', block_text)
        if len(msplit) >= 5:
            # first token may be prefix
            options = [normalize_text(x) for x in msplit[1:5]]

    # 4) Determine question text: block_text (after removing Options/Answer/Explanation pieces)
    question_text = block_text.strip()
    # remove trailing connectors like '—' at end
    question_text = re.sub(r'[\-\—\–\s]+$', '', question_text).strip()
    # If question_text still contains 'Options:' leftover remove
    question_text = re.sub(r'Options?\s*[:\-]?\s*$', '', question_text, flags=re.I).strip()

    # remove leading headings again if somehow still there
    q_lines = [ln for ln in question_text.splitlines() if ln.strip()!='']
    while q_lines and is_heading_line(q_lines[0]):
        q_lines.pop(0)
    question_text = "\n".join(q_lines).strip()

    # 5) Normalize and pad/truncate options to exactly 4
    options = [normalize_text(o) for o in options]
    # drop empty leading/trailing
    options = [o for o in options if o != "" or len(options) <= 4]  # keep empties if necessary
    if len(options) > 4:
        options = options[:4]
    while len(options) < 4:
        options.append("")

    # 6) Resolve answer if not known
    assumed = False
    if not answer_letter and raw_answer_text:
        # try to find which option text is contained in raw_answer_text (best-effort)
        ra = raw_answer_text.lower()
        for i, opt in enumerate(options):
            if opt and opt.lower() in ra:
                answer_letter = 'abcd'[i]
                break
    if not answer_letter:
        # try to check if raw answer is full text of an option
        if raw_answer_text:
            for i,opt in enumerate(options):
                if opt and normalize_text(op := opt).lower() == normalize_text(raw_answer_text).lower():
                    answer_letter = 'abcd'[i]
                    break
    if not answer_letter:
        # fallback: choose option 'a' and mark assumed
        answer_letter = 'a'
        assumed = True

    # Package parsed question
    parsed = {
        'question': normalize_text(question_text),
        'options': options,
        'answer': answer_letter,
        'assumed': assumed,
        'explanation': normalize_text(explanation),
        'raw_block': normalize_text("\n".join(paras))
    }
    # Final guard: if question_text is too short and no options and no answer, skip (likely heading)
    if (not parsed['question'] or parsed['question'].strip()=='' ) and all(o=='' for o in parsed['options']):
        return None

    return parsed

# --------------------------
# Driver parse function
# --------------------------
def parse_docx_to_questions(input_path: str, verbose: bool = True) -> List[dict]:
    paras = read_docx_paragraphs(input_path)
    blocks = group_paragraphs_into_blocks(paras)
    questions = []
    skipped = 0
    for i, block in enumerate(blocks):
        q = parse_block(block, debug_index=i)
        if q:
            questions.append(q)
        else:
            skipped += 1
            if verbose:
                # optionally show a small preview of skipped block for debugging
                preview = (" ".join(block)[:120] + "...") if block else ""
                print(f"[skip] block#{i+1} preview: {preview}")
    if verbose:
        print(f"Parsing finished: {len(questions)} questions parsed, {skipped} blocks skipped.")
    return questions

# --------------------------
# Output creation with exact 8-row table layout
# --------------------------
def create_output_docx_tables(questions: List[dict], output_path: str):
    """
    Create output .docx where each question is a separate table with exactly 8 rows and 3 columns,
    using merges per-row:
      Row0: col0 label 'Question', col1+col2 merged = question text
      Row1: col0 'Type', col1+col2 merged = 'multiple_choice'
      Row2-5: col0 'Option', col1 option text, col2 correctness ('correct'/'incorrect')
      Row6: col0 'Solution', col1+col2 merged = explanation
      Row7: col0 'Marks', col1 '1', col2 '0'
    """
    doc = Document()
    for idx, q in enumerate(questions, start=1):
        # create table with 8 rows x 3 cols
        table = doc.add_table(rows=8, cols=3)
        # set visible grid (style)
        try:
            table.style = 'Table Grid'
        except Exception:
            # style may not be available in some environments; keep default
            pass

        # Row 0: Question label + merged content
        table.cell(0,0).text = "Question"
        # merge col1 & col2 for question text
        merged_q = table.cell(0,1).merge(table.cell(0,2))
        # put question text into merged cell (preserve newlines)
        merged_q.text = q.get('question', '')

        # Row 1: Type label + merged 'multiple_choice'
        table.cell(1,0).text = "Type"
        merged_type = table.cell(1,1).merge(table.cell(1,2))
        merged_type.text = "multiple_choice"

        # Rows 2-5: options (each row is: label 'Option' | option text | correctness)
        opts = q.get('options', ['','', '',''])
        # correctness index
        letter = q.get('answer', 'a').lower()
        idx_map = {'a':0, 'b':1, 'c':2, 'd':3}
        correct_idx = idx_map.get(letter, 0)
        for r in range(2, 6):
            row_i = r
            table.cell(row_i, 0).text = "Option"
            opt_text = opts[row_i - 2] if (row_i - 2) < len(opts) else ""
            table.cell(row_i, 1).text = opt_text
            table.cell(row_i, 2).text = "correct" if (row_i - 2) == correct_idx else "incorrect"

        # Row 6: Solution label + merged explanation
        table.cell(6,0).text = "Solution"
        merged_sol = table.cell(6,1).merge(table.cell(6,2))
        merged_sol.text = q.get('explanation', '')

        # Row 7: Marks label + 1 + 0
        table.cell(7,0).text = "Marks"
        table.cell(7,1).text = "1"
        table.cell(7,2).text = "0"

        # Leave a blank paragraph between tables for readability
        doc.add_paragraph()

    # Save
    doc.save(output_path)

def main():
    root = tk.Tk()
    root.withdraw()  # hide the root window

    input_path = filedialog.askopenfilename(
        title="Select Input Word File",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not input_path:
        messagebox.showinfo("QuizFormatter", "No input file selected.")
        return

    output_path = filedialog.asksaveasfilename(
        title="Save Formatted Output As",
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not output_path:
        messagebox.showinfo("QuizFormatter", "No output file selected.")
        return

    questions = parse_docx_to_questions(input_path, verbose=False)
    if not questions:
        messagebox.showerror("QuizFormatter", "No questions could be parsed. Check your input file.")
        return

    create_output_docx_tables(questions, output_path)
    messagebox.showinfo("QuizFormatter", f"Formatted document saved at:\n{output_path}")

if __name__ == "__main__":
    main()
