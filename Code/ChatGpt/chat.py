#!/usr/bin/env python3
"""
QuizFormatter GUI

- Robust parser that preserves Word OMath (equations), Unicode/Tamil, and handles many input layouts.
- GUI with top-left app name, top-right settings (theme), centered card with Input/Output file pickers and Convert.
- Produces debug_log.jsonl with details per parsed block.

Save as quizformatter_gui.py and run:
    pip install python-docx lxml
    python quizformatter_gui.py
"""

import os
import re
import sys
import json
import unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from lxml import etree
from docx import Document

# XML namespaces for docx parsing
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

# ----------------------
# Low-level helpers
# ----------------------

def paragraph_full_text(paragraph):
    """
    Reconstruct paragraph text by iterating runs and math nodes in order.
    This preserves OMath content (extracts m:t text) and normal text.
    """
    xml = paragraph._p.xml.encode("utf-8")
    root = etree.fromstring(xml)
    pieces = []
    for child in root:
        tag = etree.QName(child.tag).localname
        if tag == "r":  # run
            for t in child.findall('.//w:t', namespaces=NSMAP):
                if t.text:
                    pieces.append(t.text)
            # preserve explicit breaks
            if child.findall('.//w:br', namespaces=NSMAP):
                pieces.append('\n')
        elif tag in ("oMath", "oMathPara"):
            # gather math text tokens
            for mt in child.findall('.//m:t', namespaces=NSMAP):
                if mt.text:
                    pieces.append(mt.text)
        else:
            # fallback: any w:t deeper in this node
            for t in child.findall('.//w:t', namespaces=NSMAP):
                if t.text:
                    pieces.append(t.text)
    return "".join([p if p is not None else "" for p in pieces]).strip()

def normalize_text(s):
    """Unicode NFC and collapse whitespace, remove zero-width chars."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFC", s)
    s = s.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
    s = re.sub(r'\r\n|\r', '\n', s)
    s = re.sub(r'\n\s+\n', '\n\n', s)
    s = re.sub(r'[ \t]+', ' ', s)
    return s.strip()

# ----------------------
# Parsing algorithm
# ----------------------

SEPARATOR_RE = re.compile(r'^[\-\—\–\*\_]{1,}\s*$')

HEADING_PATTERNS = [
    r'^unique questions',
    r'^unique questions with answers',
    r'^paper\s*\d+',
    r'^selected questions',
    r'^selected questions with answers',
    r'^questions? on',
]

def is_heading_line(line: str) -> bool:
    if not line:
        return False
    ln = line.strip().lower()
    for pat in HEADING_PATTERNS:
        if re.match(pat, ln):
            return True
    if 'unique questions' in ln or 'answers and explanations' in ln:
        return True
    if re.match(r'^paper\s*\d+\b', ln):
        return True
    return False

def group_paragraphs_into_blocks(paragraphs):
    blocks = []
    curr = []
    for p in paragraphs:
        if p is None:
            continue
        if p.strip() == "" or SEPARATOR_RE.match(p.strip()):
            if curr:
                blocks.append(curr)
                curr = []
            continue
        curr.append(p)
    if curr:
        blocks.append(curr)
    return blocks

# Core block parsing
def parse_block(block, idx):
    """
    Parse a block (list of paragraph strings) into:
    { question, options[4], answer ('a'..'d'), assumed(bool), explanation, raw_block }
    Returns None if block is considered heading/noise.
    """
    paras = [p for p in block if p and p.strip()!='']
    if not paras:
        return None
    # remove leading headings if present
    while paras and is_heading_line(paras[0]):
        paras.pop(0)
    if not paras:
        return None

    raw = " ".join(paras).strip()
    debug = {"block_index": idx, "raw_block_preview": raw[:240]}

    # extract explanation at end, if exists
    explanation = ""
    m_expl = re.search(r'(?i)\b(Explanation|Solution|Explanatory)\b\s*[:\-]?\s*(.*)$', raw, flags=re.S)
    if m_expl:
        explanation = m_expl.group(2).strip()
        raw = raw[:m_expl.start()].strip()

    # extract answer if present anywhere (prefer before options split)
    raw_answer_text = None
    answer_letter = None
    m_ans = re.search(r'(?i)\b(Answer|Ans|Correct|Key|Correct option)\b\s*[:\-]?\s*([^\n\r]*)', raw)
    if m_ans:
        raw_answer_text = m_ans.group(2).strip()
        raw = raw[:m_ans.start()].strip()
        m_letter = re.search(r'([A-Da-d])', raw_answer_text)
        if m_letter:
            answer_letter = m_letter.group(1).lower()

    # split question vs options
    question_text = raw
    options_part = ""
    m_options_token = re.search(r'(?i)\bOptions?\b\s*[:\-]?\s*(.*)$', raw, flags=re.S)
    if m_options_token:
        question_text = raw[:m_options_token.start()].strip()
        options_part = m_options_token.group(1).strip()
    else:
        # attempt split before first labeled option (a) or a. etc
        sp = re.split(r'(?=(?:\(|\[)?[A-Da-d][\)\].]?\s+)', raw, maxsplit=1)
        if len(sp) == 2:
            question_text, options_part = sp[0].strip(), sp[1].strip()
        else:
            # maybe options are on separate paragraphs - gather lines starting with a., (a) etc
            separate_opts = []
            for p in paras[1:]:
                m = re.match(r'^\s*[\(\[]?([A-Da-d])[\)\].]?\s*(.+)', p)
                if m:
                    separate_opts.append((m.group(1).lower(), m.group(2).strip()))
            if separate_opts:
                # build options_part by concatenating labeled lines
                options_part = " ".join([f"({lab}) {txt}" for lab,txt in separate_opts])
                # question_text remains the first paragraph
                question_text = paras[0].strip()

    # extract pairs label -> text using robust regex
    opt_pairs = re.findall(
        r'[\(\[]?([A-Da-d])[\)\].]?\s*'          # label
        r'([^(\(\[]+?)'                          # option text (lazy)
        r'(?=(?:[\(\[]?[A-Da-d][\)\].]?|\Z))',   # until next label or end
        options_part, flags=re.S
    )
    opts = []
    if opt_pairs:
        # sort by label order a..d
        label_map = {lab.lower(): txt.strip() for lab, txt in opt_pairs}
        for label in ['a','b','c','d']:
            opts.append(normalize_text(label_map.get(label, "")))
    else:
        # fallback: try to find inline 'a) text b) text' by splitting tokens
        inline = re.split(r'[\s]*[A-Da-d][\)\.\]]\s*', options_part)
        inline = [normalize_text(x) for x in inline if normalize_text(x)]
        if inline:
            # note: the split yields leading prefix before first label; to be safe, take last 4
            if len(inline) >= 4:
                opts = inline[:4]
            else:
                opts = inline + [""] * (4 - len(inline))
        else:
            opts = ["", "", "", ""]

    # ensure 4 options
    if len(opts) < 4:
        opts += [""] * (4 - len(opts))
    if len(opts) > 4:
        opts = opts[:4]

    # try to resolve answer by matching raw_answer_text against options if letter unknown
    assumed = False
    if not answer_letter and raw_answer_text:
        ra = raw_answer_text.lower()
        for i, o in enumerate(opts):
            if o and o.lower() in ra:
                answer_letter = 'abcd'[i]
                break

    if not answer_letter:
        # fallback choose 'a' and mark assumed
        answer_letter = 'a'
        assumed = True

    parsed = {
        'question': normalize_text(question_text),
        'options': opts,
        'answer': answer_letter,
        'assumed': assumed,
        'explanation': normalize_text(explanation),
        'raw_block': normalize_text(raw),
        'debug': {
            'block_idx': idx,
            'raw_preview': raw[:220],
            'found_options': opts,
            'raw_answer_text': raw_answer_text,
            'final_answer': answer_letter,
            'assumed': assumed
        }
    }

    # If parsed question empty and options empty treat as not a question
    if (not parsed['question']) and all(o == "" for o in parsed['options']):
        return None
    return parsed

def parse_docx_to_questions(input_path, write_debug_log=True):
    """
    Parses the input docx and returns a list of parsed question dicts.
    Also writes debug_log.jsonl if write_debug_log True.
    """
    doc = Document(input_path)
    paragraphs = [paragraph_full_text(p) for p in doc.paragraphs]
    blocks = group_paragraphs_into_blocks(paragraphs)
    questions = []
    debug_entries = []
    for i, block in enumerate(blocks):
        q = parse_block(block, i)
        if q:
            questions.append(q)
            debug_entries.append(q['debug'])
    if write_debug_log:
        try:
            with open("debug_log.jsonl", "w", encoding="utf-8") as f:
                for e in debug_entries:
                    f.write(json.dumps(e, ensure_ascii=False) + "\n")
        except Exception as e:
            # ignore write debug errors; not critical
            pass
    return questions

# ----------------------
# Output builder
# ----------------------

def write_output_docx(questions, output_path):
    """
    Write the exact required table per question:
    8 rows x 3 cols:
    Row0: cell(0,0) = 'Question', cell(0,1)+cell(0,2) merged -> question text
    Row1: cell(1,0) = 'Type', merge cell(1,1..2) -> 'multiple_choice'
    Row2-5: each row: 'Option' | option text | correct/incorrect
    Row6: 'Solution' | merge cell(6,1..2) -> explanation
    Row7: 'Marks' | '1' | '0'
    """
    doc = Document()
    for q in questions:
        table = doc.add_table(rows=8, cols=3)
        try:
            table.style = 'Table Grid'
        except Exception:
            pass

        # Row0
        table.cell(0,0).text = "Question"
        table.cell(0,1).merge(table.cell(0,2)).text = q.get('question', '')

        # Row1
        table.cell(1,0).text = "Type"
        table.cell(1,1).merge(table.cell(1,2)).text = "multiple_choice"

        # Rows 2-5: Options
        opts = q.get('options', ['','','',''])
        letter = q.get('answer','a').lower()
        idx_map = {'a':0,'b':1,'c':2,'d':3}
        correct_idx = idx_map.get(letter, 0)
        for i in range(4):
            r = 2 + i
            table.cell(r,0).text = "Option"
            table.cell(r,1).text = opts[i] or ""
            table.cell(r,2).text = "correct" if i==correct_idx else "incorrect"

        # Row6 solution
        table.cell(6,0).text = "Solution"
        table.cell(6,1).merge(table.cell(6,2)).text = q.get('explanation','')

        # Row7 marks
        table.cell(7,0).text = "Marks"
        table.cell(7,1).text = "1"
        table.cell(7,2).text = "0"

        # spacing paragraph
        doc.add_paragraph()
    doc.save(output_path)

# ----------------------
# GUI
# ----------------------

class QuizFormatterGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QuizFormatter")
        self.geometry("980x700")
        self.minsize(880, 620)

        # theme state
        self.theme = tk.StringVar(value="light")

        # file paths
        self.input_path = tk.StringVar(value="")
        self.output_path = tk.StringVar(value="")

        # parsed questions cache
        self.questions = []

        # build UI
        self._build_style()
        self._build_top_bar()
        self._build_center_card()
        self._build_bottom_area()

        # apply initial theme
        self.apply_theme(self.theme.get())

    def _build_style(self):
        self.style = ttk.Style(self)
        try:
            # prefer clam for consistent look
            self.style.theme_use("clam")
        except Exception:
            pass
        # custom styles
        self.style.configure("App.TFrame", background="#f3f6f9")
        self.style.configure("Card.TFrame", background="#ffffff", relief="flat")
        self.style.configure("Header.TLabel", font=("Inter", 18, "bold"), background="#f3f6f9")
        self.style.configure("Muted.TLabel", font=("Inter", 10), foreground="#666666", background="#f3f6f9")
        self.style.configure("TButton", padding=6)

    def _build_top_bar(self):
        top = ttk.Frame(self, style="App.TFrame", padding=(12,10))
        top.pack(side=tk.TOP, fill=tk.X)
        # App name top-left
        lbl = ttk.Label(top, text="QuizFormatter", style="Header.TLabel")
        lbl.pack(side=tk.LEFT)

        # Spacer
        top_spacer = ttk.Frame(top)
        top_spacer.pack(side=tk.LEFT, expand=True)

        # Settings gear top-right (button)
        gear_btn = ttk.Button(top, text="⚙️ Settings", command=self.open_settings)
        gear_btn.pack(side=tk.RIGHT)

    def _build_center_card(self):
        # center frame
        center = ttk.Frame(self, style="App.TFrame")
        center.pack(fill=tk.BOTH, expand=True, padx=24, pady=10)

        # Card frame (centered)
        card = ttk.Frame(center, style="Card.TFrame", padding=20)
        card.place(relx=0.5, rely=0.5, anchor=tk.CENTER, relwidth=0.76, relheight=0.6)

        # Title (inside card)
        title = ttk.Label(card, text="Convert exam questions to professor format", font=("Inter", 14, "bold"))
        title.pack(anchor=tk.W, pady=(0,8))

        desc = ttk.Label(card, text="Select the input .docx file and choose where to save the formatted output.\nThe parser preserves math, Tamil/Unicode, and writes exact 8-row tables per question.",
                         style="Muted.TLabel", justify=tk.LEFT)
        desc.pack(anchor=tk.W, pady=(0,12))

        # Form-like area
        form = ttk.Frame(card)
        form.pack(fill=tk.BOTH, expand=True)

        # Input row
        in_row = ttk.Frame(form)
        in_row.pack(fill=tk.X, pady=8)
        in_label = ttk.Label(in_row, text="Input (.docx):", width=12)
        in_label.pack(side=tk.LEFT, padx=(0,6))

        # use tk.Entry for easier bg/fg styling across themes
        self.input_entry = tk.Entry(in_row, textvariable=self.input_path, font=("Inter", 11))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,8))

        browse_btn = ttk.Button(in_row, text="Browse", command=self.browse_input)
        browse_btn.pack(side=tk.LEFT)

        # Output row
        out_row = ttk.Frame(form)
        out_row.pack(fill=tk.X, pady=8)
        out_label = ttk.Label(out_row, text="Output (.docx):", width=12)
        out_label.pack(side=tk.LEFT, padx=(0,6))

        self.output_entry = tk.Entry(out_row, textvariable=self.output_path, font=("Inter", 11))
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,8))

        save_btn = ttk.Button(out_row, text="Save As", command=self.save_as)
        save_btn.pack(side=tk.LEFT)

        # Convert button (centered)
        c_row = ttk.Frame(form)
        c_row.pack(fill=tk.X, pady=(18,0))
        convert_btn = ttk.Button(c_row, text="Convert", command=self.on_convert)
        convert_btn.pack(side=tk.TOP, pady=(0,6), ipadx=12)

        # Preview & notes
        preview_label = ttk.Label(form, text="Preview (parsed summary):", style="Muted.TLabel")
        preview_label.pack(anchor=tk.W, pady=(14,0))

        self.preview_box = ScrolledText(form, height=6)
        self.preview_box.pack(fill=tk.BOTH, expand=True, pady=(4,0))

    def _build_bottom_area(self):
        bottom = ttk.Frame(self, style="App.TFrame", padding=8)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        log_label = ttk.Label(bottom, text="Activity Log:", style="Muted.TLabel")
        log_label.pack(anchor=tk.W)
        self.log_box = ScrolledText(bottom, height=8)
        self.log_box.pack(fill=tk.BOTH, expand=True, pady=(4,0))

        # set read-only for log and preview
        self.preview_box.config(state=tk.DISABLED)
        self.log_box.config(state=tk.DISABLED)

    # ------------------------
    # UI actions
    # ------------------------
    def browse_input(self):
        p = filedialog.askopenfilename(title="Select input .docx file", filetypes=[("Word Documents", "*.docx")])
        if p:
            self.input_path.set(p)
            # set a default output filename next to input
            base = os.path.splitext(os.path.basename(p))[0]
            suggested = os.path.join(os.path.dirname(p), base + "_Formatted.docx")
            if not self.output_path.get():
                self.output_path.set(suggested)
            self.log(f"Selected input: {p}")

    def save_as(self):
        default = self.output_path.get() or ""
        initialdir = os.path.dirname(default) if default else os.getcwd()
        savep = filedialog.asksaveasfilename(title="Save formatted output as",
                                             defaultextension=".docx",
                                             filetypes=[("Word Documents", "*.docx")],
                                             initialdir=initialdir,
                                             initialfile=os.path.basename(default) if default else "")
        if savep:
            self.output_path.set(savep)
            self.log(f"Selected output: {savep}")

    def on_convert(self):
        inp = self.input_path.get().strip()
        outp = self.output_path.get().strip()
        if not inp or not os.path.isfile(inp):
            messagebox.showerror("Input missing", "Please choose a valid input .docx file.")
            return
        if not outp:
            messagebox.showerror("Output missing", "Please choose where to save the output .docx file.")
            return

        self.log("Parsing input...")
        try:
            questions = parse_docx_to_questions(inp, write_debug_log=True)
        except Exception as e:
            self.log(f"Error while parsing: {e}")
            messagebox.showerror("Parse error", f"Failed to parse input file:\n{e}")
            return

        self.questions = questions
        self.log(f"Parsed {len(questions)} question(s). Debug log written to debug_log.jsonl (in current folder).")
        # preview first few
        self.preview_box.config(state=tk.NORMAL)
        self.preview_box.delete("1.0", tk.END)
        if questions:
            for i, q in enumerate(questions[:5]):
                self.preview_box.insert(tk.END, f"Q{i+1}: {q['question'][:200]}\n")
                for j,opt in enumerate(q['options']):
                    label = ['A','B','C','D'][j]
                    corr = " (correct)" if q['answer']==['a','b','c','d'][j] else ""
                    self.preview_box.insert(tk.END, f"  {label}. {opt}{corr}\n")
                if q['explanation']:
                    self.preview_box.insert(tk.END, f"  Solution: {q['explanation']}\n")
                if q.get('assumed'):
                    self.preview_box.insert(tk.END, "  [Assumed answer: A]\n")
                self.preview_box.insert(tk.END, "-"*40 + "\n")
        else:
            self.preview_box.insert(tk.END, "No questions detected. Check input file formatting.")

        self.preview_box.config(state=tk.DISABLED)

        # Ask user to confirm export if there are assumed answers or missing options
        warnings = []
        for i,q in enumerate(questions):
            if q.get('assumed'):
                warnings.append(f"Q{i+1}: assumed answer (A).")
            if any(o.strip()=="" for o in q['options']):
                warnings.append(f"Q{i+1}: fewer than 4 options (padded empty).")
        if warnings:
            proceed = messagebox.askyesno("Parsing Warnings",
                                          "Parser made the following assumptions or found issues:\n\n" +
                                          "\n".join(warnings[:10]) +
                                          ("\n\n(Showing first 10)\n\nProceed with export?"))
            if not proceed:
                self.log("Export canceled by user due to warnings.")
                return

        # write output
        try:
            write_output_docx(questions, outp)
            self.log(f"Exported formatted docx to: {outp}")
            messagebox.showinfo("Success", f"Formatted document saved:\n{outp}")
        except Exception as e:
            self.log(f"Failed to save output: {e}")
            messagebox.showerror("Save error", f"Failed to write output file:\n{e}")

    # ------------------------
    # Logging & settings
    # ------------------------
    def log(self, text):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Settings")
        win.geometry("340x160")
        win.transient(self)
        ttk.Label(win, text="Theme:").pack(anchor=tk.W, padx=12, pady=(10,4))
        frame = ttk.Frame(win)
        frame.pack(anchor=tk.W, padx=12)
        ttk.Radiobutton(frame, text="Light", variable=self.theme, value="light", command=lambda: self.apply_theme("light")).pack(anchor=tk.W)
        ttk.Radiobutton(frame, text="Dark", variable=self.theme, value="dark", command=lambda: self.apply_theme("dark")).pack(anchor=tk.W)
        ttk.Label(win, text="(Theme affects background & editor colors)", style="Muted.TLabel").pack(anchor=tk.W, padx=12, pady=(8,0))

    def apply_theme(self, theme_name):
        """Apply a simple light/dark color scheme."""
        theme_name = theme_name or "light"
        bg_light = "#f3f6f9"
        card_light = "#ffffff"
        fg_light = "#111111"

        bg_dark = "#1f2226"
        card_dark = "#2b2f33"
        fg_dark = "#e8eef6"

        if theme_name == "dark":
            bg = bg_dark; card = card_dark; fg = fg_dark
            entry_bg = "#3a3f44"
            entry_fg = "#ffffff"
            text_bg = "#25282b"
            text_fg = "#e8eef6"
        else:
            bg = bg_light; card = card_light; fg = fg_light
            entry_bg = "#ffffff"
            entry_fg = "#000000"
            text_bg = "#ffffff"
            text_fg = "#111111"

        # root bg
        self.configure(background=bg)
        # style frames
        self.style.configure("App.TFrame", background=bg)
        self.style.configure("Card.TFrame", background=card)
        self.style.configure("Header.TLabel", background=bg)
        self.style.configure("Muted.TLabel", background=bg, foreground=("#bdbdbd" if theme_name=="dark" else "#666666"))
        # entries are tk.Entry -> set directly
        try:
            self.input_entry.config(bg=entry_bg, fg=entry_fg, insertbackground=entry_fg)
            self.output_entry.config(bg=entry_bg, fg=entry_fg, insertbackground=entry_fg)
            self.preview_box.config(bg=text_bg, fg=text_fg, insertbackground=text_fg)
            self.log_box.config(bg=text_bg, fg=text_fg, insertbackground=text_fg)
        except Exception:
            pass

# ----------------------
# Program entry
# ----------------------

def main():
    # dependency checks
    try:
        import docx  # noqa: F401
        from lxml import etree  # noqa: F401
    except Exception as e:
        message = "Missing dependencies: please run\n\n    pip install python-docx lxml\n\nThen re-run this program."
        print(message)
        messagebox.showerror("Missing dependencies", message)
        return

    app = QuizFormatterGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
