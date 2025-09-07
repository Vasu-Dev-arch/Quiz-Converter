import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from docx import Document
import re
import os

def merge_cells(row, start_idx, end_idx):
    """Merge cells horizontally from start_idx to end_idx in given row."""
    row.cells[start_idx].merge(row.cells[end_idx])

def create_question_table(doc, qdata):
    """Create a fully merged, gridlined table per question."""

    table = doc.add_table(rows=8, cols=3)
    table.style = 'Table Grid'

    # Row 0: Question (merge col 1 & 2)
    table.cell(0, 0).text = "Question"
    qcell = table.cell(0, 1)
    qcell.text = qdata['question']
    merge_cells(table.rows[0], 1, 2)

    # Row 1: Type (merge col 1 & 2)
    table.cell(1, 0).text = "Type"
    tcell = table.cell(1, 1)
    tcell.text = "multiple_choice"
    merge_cells(table.rows[1], 1, 2)

    # Rows 2-5: Options with correctness
    for i in range(4):
        table.cell(2 + i, 0).text = "Option"
        table.cell(2 + i, 1).text = qdata['options'][i]
        table.cell(2 + i, 2).text = "correct" if i == qdata['correct'] else "incorrect"

    # Row 6: Solution (merge col 1 & 2)
    table.cell(6, 0).text = "Solution"
    scell = table.cell(6,1)
    scell.text = qdata.get('solution', '')
    merge_cells(table.rows[6], 1, 2)

    # Row 7: Marks
    table.cell(7, 0).text = "Marks"
    table.cell(7, 1).text = "1"
    table.cell(7, 2).text = "0"

def find_answer_index(options, ans_text):
    """Return option index matching answer letter or text."""
    ans_text = ans_text.strip().lower()
    # Extract letter from answer if present
    letter_match = re.match(r'^[\(\[]?([a-d])[)\]]?', ans_text)
    if letter_match:
        idx = ord(letter_match.group(1)) - ord('a')
        if 0 <= idx < 4:
            return idx
    # Search in options for answer text fragment matching
    for i, opt in enumerate(options):
        if ans_text in opt.lower():
            return i
    return 0  # fallback

def parse_docx_questions(filepath):
    """Parse questions, options, answers, solutions from given docx file."""

    doc = Document(filepath)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(paras)

    # Split questions by dashes or numbered pattern
    raw_blocks = re.split(r'\n\s*(?:---|—)\s*\n', full_text)
    questions = []

    for block in raw_blocks:
        if not block.strip():
            continue
        if any(x in block.lower() for x in ['title', 'paper', 'unique questions']):
            continue

        lines = [line.strip() for line in block.split('\n') if line.strip()]
        question_text = ""
        options = []
        answer = ""
        solution = ""

        # Parsing state machine approach
        stage = 'question'
        for line in lines:
            if re.match(r"Options?:", line, re.IGNORECASE):
                stage = 'options'
                # Extract inline options if present e.g: Options: (a) xxx (b) yyy ...
                opts_inline = re.findall(r'\(?([a-d])\)?[\.\)]?\s*([^(\[]+?)(?=\s*\(?[a-d]\)?[\.\)]|$)', line, re.I)
                if opts_inline:
                    options = [opt[1].strip() for opt in opts_inline]
                continue
            elif re.match(r'^(Answer|Ans|Correct):', line, re.IGNORECASE):
                stage = 'answer'
                answer = re.sub(r'^(Answer|Ans|Correct):\s*', '', line, flags=re.I).strip()
                continue
            elif re.match(r'^(Explanation|Solution):', line, re.IGNORECASE):
                stage = 'solution'
                solution = re.sub(r'^(Explanation|Solution):\s*', '', line, flags=re.I).strip()
                continue

            # Accumulate text depending on stage
            if stage == 'question':
                if not question_text:
                    question_text = line
                else:
                    question_text += " " + line
            elif stage == 'options':
                # Collect line-by-line options, e.g. "a. text", "(a) text"
                match = re.match(r'\(?([a-d])\)?[\.\)]\s*(.*)', line, re.I)
                if match:
                    options.append(match.group(2).strip())
                else:
                    # If options inline not found, treat lines directly as options if fewer than 4
                    if len(options) < 4:
                        options.append(line)
            elif stage == 'solution':
                solution += (" " + line).strip()

        # Ensure 4 options, pad empty strings if fewer
        while len(options) < 4:
            options.append("")

        # Determine correct answer index
        correct_idx = find_answer_index(options, answer) if answer else 0

        questions.append({
            'question': question_text,
            'options': options,
            'correct': correct_idx,
            'solution': solution,
        })

    return questions


class QuizFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QuizFormatter - Professional Quiz Converter")

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Accent.TButton", foreground="#fff", background="#0066CC", font=('Segoe UI', 11, "bold"))
        style.map("Accent.TButton", background=[('active', '#005bb5')])
        self.root.config(bg="#f5f7fa")

        # Header
        ttk.Label(root, text="QuizFormatter", font=("Segoe UI", 24, "bold"), foreground="#0066CC", background="#f5f7fa").pack(pady=(30,10))
        ttk.Label(root, text="Convert Quiz/Exam Questions Word files to Structured Tables", font=("Segoe UI", 11), foreground="#333", background="#f5f7fa").pack()

        # Main frame
        main_frame = ttk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=30, pady=15)

        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill='x')

        self.file_label = ttk.Label(file_frame, text="No file selected", font=("Segoe UI", 9), foreground="#555")
        self.file_label.pack(side="left")

        self.upload_btn = ttk.Button(file_frame, text="Upload Input File (.docx)", style="Accent.TButton", command=self.upload_file)
        self.upload_btn.pack(side="right")

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15, fill='x')

        self.preview_btn = ttk.Button(btn_frame, text="Preview Questions", state='disabled', command=self.preview_questions)
        self.preview_btn.pack(side="left", padx=5)

        self.convert_btn = ttk.Button(btn_frame, text="Convert & Save", state='disabled', command=self.convert_save)
        self.convert_btn.pack(side="left", padx=5)

        # Status console
        console_frame = ttk.LabelFrame(main_frame, text="Status Log")
        console_frame.pack(fill="both", expand=True)

        self.console = ScrolledText(console_frame, height=12, font=("Consolas", 10))
        self.console.pack(fill="both", expand=True)

        self.input_file = None
        self.questions = []

    def log(self, message, tag=None):
        self.console.insert('end', message + '\n')
        if tag:
            self.console.tag_add(tag, "end-1c linestart", "end-1c lineend")
            if tag == "error":
                self.console.tag_configure(tag, foreground="red")
            elif tag == "success":
                self.console.tag_configure(tag, foreground="green")
            elif tag == "warning":
                self.console.tag_configure(tag, foreground="orange")
        self.console.see('end')

    def upload_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
        if not filepath:
            return
        self.input_file = filepath
        self.file_label.config(text=os.path.basename(filepath))
        self.log(f"File selected: {filepath}", "success")

        try:
            self.questions = parse_docx_questions(filepath)
            if not self.questions:
                self.log("Warning: No questions detected in file.", "warning")
                messagebox.showwarning("No Questions Found", "No questions were detected in the uploaded file. Please check file formatting.")
                self.preview_btn.config(state='disabled')
                self.convert_btn.config(state='disabled')
            else:
                self.log(f"{len(self.questions)} questions parsed successfully.", "success")
                self.preview_btn.config(state='enabled')
                self.convert_btn.config(state='enabled')
        except Exception as e:
            self.log(f"Error parsing file: {e}", "error")
            messagebox.showerror("Parsing Error", f"An error occurred while parsing the file:\n{e}")
            self.preview_btn.config(state='disabled')
            self.convert_btn.config(state='disabled')

    def preview_questions(self):
        if not self.questions:
            messagebox.showinfo("No Data", "There are no parsed questions to preview.")
            return
        preview_win = tk.Toplevel(self.root)
        preview_win.title("Question Preview")
        preview_win.geometry("900x600")

        txt_widget = ScrolledText(preview_win, font=("Segoe UI", 10))
        txt_widget.pack(fill="both", expand=True)

        for idx, q in enumerate(self.questions, 1):
            txt_widget.insert('end', f"Question {idx}:\n{q['question']}\nOptions:\n")
            for i, opt in enumerate(q['options']):
                mark = "(correct)" if i == q['correct'] else ""
                txt_widget.insert('end', f"  {chr(97+i)}. {opt} {mark}\n")
            txt_widget.insert('end', f"Solution: {q['solution']}\n\n{'-'*60}\n\n")
        txt_widget.configure(state='disabled')

    def convert_save(self):
        if not self.questions:
            messagebox.showerror("Error", "No questions to convert. Please upload and parse a file first.")
            return
        outpath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
        if not outpath:
            return

        outdoc = Document()
        try:
            for qdata in self.questions:
                create_question_table(outdoc, qdata)
                outdoc.add_paragraph("")

            outdoc.save(outpath)
            self.log(f"Output saved successfully: {outpath}", "success")
            messagebox.showinfo("Success", f"Questions converted and saved to:\n{outpath}")
        except Exception as e:
            self.log(f"Error saving file: {e}", "error")
            messagebox.showerror("Save Error", f"Failed to save output file:\n{e}")

def main():
    root = tk.Tk()
    app = QuizFormatterApp(root)
    root.mainloop()

if __name__=="__main__":
    main()

