import re
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Inches

def parse_questions(doc):
    # Join all non-empty paragraph texts with newline
    full_text = '\n'.join(para.text for para in doc.paragraphs if para.text.strip())
    
    # Split into question blocks using the separator '—'
    blocks = re.split(r'—', full_text)
    
    questions = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        
        # Split block into lines
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if not lines:
            continue
        
        # Parse first line: assuming format like "1. Question text: Options: (a) opt1 (b) opt2 ..."
        first_line = lines[0]
        match = re.match(r'^\d+\.\s*(.+?)\s*Options:\s*(.+)$', first_line, re.IGNORECASE)
        if match:
            question = match.group(1).strip()
            opts_str = match.group(2).strip()
            
            # Parse options using regex to find (a) text, (b) text, etc.
            option_matches = re.findall(r'\(([a-d])\)\s*([^()]+?)(?=\s*\([a-d]\)|$)', opts_str, re.IGNORECASE)
            options = [opt.strip() for _, opt in option_matches]
        else:
            continue  # Not a valid question block
        
        # Pad or truncate options to exactly 4
        while len(options) < 4:
            options.append('')
        options = options[:4]
        
        # Default correct index to 0 if no answer found
        correct_index = 0
        solution = ''
        
        # Parse remaining lines for answer and explanation
        i = 1
        while i < len(lines):
            line = lines[i]
            if re.match(r'^Answer:\s*', line, re.IGNORECASE):
                ans_str = re.split(r'^Answer:\s*', line, flags=re.IGNORECASE)[1].strip()
                letter_match = re.match(r'\(([a-d])\)', ans_str, re.IGNORECASE)
                if letter_match:
                    correct_index = ord(letter_match.group(1).lower()) - ord('a')
            elif re.match(r'^Explanation:\s*', line, re.IGNORECASE):
                sol_str = re.split(r'^Explanation:\s*', line, flags=re.IGNORECASE)[1].strip()
                solution += sol_str + '\n'
                # Collect subsequent lines as part of solution until end
                i += 1
                while i < len(lines):
                    solution += lines[i] + '\n'
                    i += 1
                break
            i += 1
        
        questions.append({
            'question': question,
            'options': options,
            'correct_index': correct_index,
            'solution': solution.strip()
        })
    
    return questions

def create_output_table(output_doc, q):
    # Add a table with 8 rows and 3 columns
    table = output_doc.add_table(rows=8, cols=3)
    table.style = 'Table Grid'
    table.autofit = False
    # Set column widths (optional, for better appearance)
    table.columns[0].width = Inches(1.0)
    table.columns[1].width = Inches(4.5)
    table.columns[2].width = Inches(1.0)
    
    # Row 0: Question
    table.rows[0].cells[0].text = 'Question'
    table.rows[0].cells[1].text = q['question']
    table.rows[0].cells[2].text = ''
    
    # Row 1: Type
    table.rows[1].cells[0].text = 'Type'
    table.rows[1].cells[1].text = 'multiple_choice'
    table.rows[1].cells[2].text = ''
    
    # Rows 2-5: Options
    for i in range(4):
        row_idx = 2 + i
        table.rows[row_idx].cells[0].text = 'Option'
        table.rows[row_idx].cells[1].text = q['options'][i]
        correctness = 'correct' if i == q['correct_index'] else 'incorrect'
        table.rows[row_idx].cells[2].text = correctness
    
    # Row 6: Solution
    sol_cell = table.rows[6].cells[0].text = 'Solution'
    table.rows[6].cells[2].text = ''
    sol_content_cell = table.rows[6].cells[1]
    # Handle multi-line solution by adding paragraphs
    lines = [line.strip() for line in q['solution'].split('\n') if line.strip()]
    if lines:
        sol_content_cell.paragraphs[0].text = lines[0]
        for line in lines[1:]:
            sol_content_cell.add_paragraph(line)
    else:
        sol_content_cell.text = ''
    
    # Row 7: Marks
    table.rows[7].cells[0].text = 'Marks'
    table.rows[7].cells[1].text = '1'
    table.rows[7].cells[2].text = '0'
    
    # Add spacing after each table
    output_doc.add_paragraph()

def upload_file():
    path = filedialog.askopenfilename(filetypes=[('Word files', '*.docx')])
    if path:
        input_path.set(path)
        status.set(f'File uploaded: {path}')

def convert_and_save():
    in_path = input_path.get()
    if not in_path:
        messagebox.showerror('Error', 'No input file selected.')
        return
    
    out_path = filedialog.asksaveasfilename(defaultextension='.docx', filetypes=[('Word files', '*.docx')])
    if not out_path:
        return
    
    try:
        input_doc = Document(in_path)
        questions = parse_questions(input_doc)
        
        if not questions:
            messagebox.showwarning('Warning', 'No questions detected in the input file.')
            return
        
        output_doc = Document()
        for q in questions:
            create_output_table(output_doc, q)
        
        output_doc.save(out_path)
        status.set(f'Converted and saved to: {out_path}')
    except Exception as e:
        messagebox.showerror('Error', f'An error occurred: {str(e)}')
        status.set(f'Error: {str(e)}')

# GUI Setup
root = tk.Tk()
root.title('QuizFormatter')
root.geometry('600x300')

input_path = tk.StringVar()
status = tk.StringVar(value='Ready')

tk.Label(root, text='QuizFormatter', font=('Arial', 16)).pack(pady=10)

tk.Button(root, text='Upload Input File (.docx)', command=upload_file).pack(pady=5)
tk.Entry(root, textvariable=input_path, width=70).pack(pady=5)

tk.Button(root, text='Convert & Save', command=convert_and_save).pack(pady=10)

tk.Label(root, textvariable=status, wraplength=500).pack(pady=10)

root.mainloop()