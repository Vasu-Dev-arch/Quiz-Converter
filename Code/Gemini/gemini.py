import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Inches
import os
import re

class Question:
    """A class to represent a single quiz question."""
    def __init__(self, question_text, options, correct_option, explanation=""):
        self.question_text = question_text
        self.options = options
        self.correct_option = correct_option
        self.explanation = explanation

def parse_docx(filepath):
    """
    Parses a .docx file and extracts questions, options, and answers.

    Args:
        filepath (str): The path to the input .docx file.

    Returns:
        list: A list of Question objects.
    """
    questions = []
    try:
        doc = Document(filepath)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open the document: {e}")
        return questions

    full_text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
    
    # Split the document content by a delimiter that separates questions
    # Common delimiters are "—" or a new question starting with an uppercase letter
    # This regex looks for a line starting with an uppercase letter or an underscore followed by a newline.
    # We will use this to split the document into potential questions.
    # The uploaded file uses "—" as a separator, so let's handle that.
    
    raw_questions = re.split(r'—\s*', full_text)
    
    if not raw_questions or all(not q.strip() for q in raw_questions):
        messagebox.showwarning("Warning", "No questions detected in the document.")
        return questions

    for raw_q in raw_questions:
        raw_q = raw_q.strip()
        if not raw_q:
            continue

        lines = raw_q.split('\n')
        question_text_lines = []
        options = []
        correct_option_raw = None
        explanation = ""
        
        state = "QUESTION"
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # Check for options
            option_match = re.match(r'^(?:[a-d]|\([a-d]\))\.\s*(.*)', line, re.IGNORECASE)
            
            # Look for answer line
            answer_match = re.match(r'^(?:Answer:|Ans:|Correct:|Correct option:)\s*(.*)', line, re.IGNORECASE)
            
            # Look for explanation line
            explanation_match = re.match(r'^Explanation:\s*(.*)', line, re.IGNORECASE)

            if state == "QUESTION":
                if option_match:
                    options.append(option_match.group(1).strip())
                    state = "OPTIONS"
                elif answer_match:
                    correct_option_raw = answer_match.group(1).strip()
                    state = "ANSWER"
                else:
                    question_text_lines.append(line)
            
            elif state == "OPTIONS":
                if answer_match:
                    correct_option_raw = answer_match.group(1).strip()
                    state = "ANSWER"
                elif explanation_match:
                    explanation += explanation_match.group(1).strip()
                    state = "EXPLANATION"
                elif option_match:
                    options.append(option_match.group(1).strip())
                elif not question_text_lines and len(options) == 0:
                    # In case there's no explicit 'a.' or '(a)' prefix, assume the next lines are options
                    options.append(line)
                else:
                    # If we are in the options state, and the next line doesn't look like an option
                    # or an answer, we assume it's part of the last option or question.
                    pass

            elif state == "ANSWER":
                if explanation_match:
                    explanation += explanation_match.group(1).strip()
                    state = "EXPLANATION"
                else:
                    # The rest of the text after the answer and before the next question is the explanation
                    explanation += " " + line
            
            elif state == "EXPLANATION":
                explanation += " " + line

        question_text = " ".join(question_text_lines).strip()
        explanation = explanation.strip()

        # Handle cases where options don't have a prefix, but are just a list of lines.
        # This is a bit tricky, so let's check if the raw question has "Options:"
        # The uploaded file uses "Options: (a) ... " format.
        if "Options:" in raw_q:
            option_part = raw_q.split("Options:")[1].split("Answer:")[0].strip()
            # Find all options and their corresponding text
            options_with_text = re.findall(r'\s*\([a-d]\)\s*(.*?)(?=\s*\([a-d]\)|\s*$)', option_part)
            options = [o.strip() for o in options_with_text if o.strip()]

        # Determine the correct option
        correct_option_text = None
        if correct_option_raw:
            # Try to match by letter first
            match_letter = re.match(r'^\s*\(?([a-d])\)?\s*$', correct_option_raw, re.IGNORECASE)
            if match_letter:
                correct_option_letter = match_letter.group(1).lower()
                if len(options) > 0:
                    correct_option_text = options[ord(correct_option_letter) - ord('a')]
            
            # If not by letter, try to match by text
            if not correct_option_text:
                for opt in options:
                    if opt.strip().lower() == correct_option_raw.lower().replace(f'({correct_option_raw.lower()})', '').strip():
                        correct_option_text = opt
                        break

        # If no explicit answer found, mark the first option as correct
        if not correct_option_text and len(options) > 0:
            correct_option_text = options[0]

        # Ensure exactly 4 options
        if len(options) < 4:
            options.extend([""] * (4 - len(options)))
        elif len(options) > 4:
            options = options[:4]
            
        questions.append(Question(question_text, options, correct_option_text, explanation))
            
    return questions

def generate_output_doc(questions, output_filepath):
    """
    Generates the output .docx document with formatted tables.

    Args:
        questions (list): A list of Question objects.
        output_filepath (str): The path to save the output .docx file.
    """
    if not questions:
        messagebox.showinfo("Info", "No questions to convert. Aborting conversion.")
        return

    doc = Document()
    
    # Headers for the table
    headers = ['Question', 'Type', 'Option', 'Option', 'Option', 'Option', 'Solution', 'Marks']
    
    for q in questions:
        table = doc.add_table(rows=2, cols=8, style='Table Grid')
        
        # Add header row
        for i, header_text in enumerate(headers):
            table.cell(0, i).text = header_text
            
        # Add content row
        table.cell(1, 0).text = q.question_text
        table.cell(1, 1).text = "multiple_choice"
        
        # Add options with correctness
        for i, option_text in enumerate(q.options):
            correctness = "correct" if option_text == q.correct_option else "incorrect"
            table.cell(1, 2 + i).text = f"{option_text} {correctness}"
            
        table.cell(1, 6).text = q.explanation
        table.cell(1, 7).text = "1 0"
        
        # Add a paragraph to separate the tables
        doc.add_paragraph()

    try:
        doc.save(output_filepath)
        messagebox.showinfo("Success", f"Conversion complete! File saved at:\n{output_filepath}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save the document: {e}")

class QuizFormatterApp(tk.Tk):
    """The main application class for QuizFormatter."""
    def __init__(self):
        super().__init__()
        self.title("QuizFormatter")
        self.geometry("600x400")
        
        self.input_filepath = ""

        self.create_widgets()
        
    def create_widgets(self):
        """Creates the GUI widgets for the application."""
        # Header
        header_label = tk.Label(self, text="QuizFormatter", font=("Helvetica", 24, "bold"))
        header_label.pack(pady=20)
        
        # Frame for buttons
        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        
        self.upload_button = tk.Button(button_frame, text="Upload Input File (.docx)", command=self.upload_file)
        self.upload_button.pack(side=tk.LEFT, padx=10)
        
        self.convert_button = tk.Button(button_frame, text="Convert & Save", command=self.convert_file, state=tk.DISABLED)
        self.convert_button.pack(side=tk.LEFT, padx=10)
        
        # Status log
        self.status_label = tk.Label(self, text="Status: Ready", relief=tk.SUNKEN, bd=1, anchor="w", fg="blue")
        self.status_label.pack(fill=tk.X, pady=10, padx=10)

    def upload_file(self):
        """Opens a file dialog for the user to select an input .docx file."""
        self.input_filepath = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")]
        )
        if self.input_filepath:
            self.status_label.config(text=f"Status: File selected: {os.path.basename(self.input_filepath)}", fg="green")
            self.convert_button.config(state=tk.NORMAL)
        else:
            self.status_label.config(text="Status: File selection cancelled.", fg="red")
            self.convert_button.config(state=tk.DISABLED)

    def convert_file(self):
        """Parses the uploaded file and saves the formatted output."""
        if not self.input_filepath:
            messagebox.showerror("Error", "Please upload an input file first.")
            return

        self.status_label.config(text="Status: Parsing and converting...", fg="orange")
        self.update_idletasks()
        
        try:
            questions = parse_docx(self.input_filepath)
            
            if not questions:
                self.status_label.config(text="Status: Conversion failed. No valid questions found.", fg="red")
                return

            output_filepath = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Documents", "*.docx")],
                initialfile=f"Formatted_Quiz_{os.path.basename(self.input_filepath)}"
            )
            
            if output_filepath:
                generate_output_doc(questions, output_filepath)
                self.status_label.config(text=f"Status: Conversion successful.", fg="green")
            else:
                self.status_label.config(text="Status: Save cancelled.", fg="red")
        except Exception as e:
            self.status_label.config(text=f"Status: An error occurred during conversion.", fg="red")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    app = QuizFormatterApp()
    app.mainloop()
