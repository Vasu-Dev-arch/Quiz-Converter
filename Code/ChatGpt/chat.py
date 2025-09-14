#!/usr/bin/env python3
"""
QuizFormatter (fixed version)

- Reads a .docx with mixed-format questions (English/Tamil, inline options, math).
- Produces output .docx with one table per question, exactly 8 rows, "Table Grid" style.
- Preserves Unicode and Word OMath equations.
- Logs parsed blocks into debug_log.jsonl for review.
"""

import sys, os, re, json, unicodedata
from typing import List, Optional
from lxml import etree
from docx import Document

# Namespaces for XML parsing
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

# ------------------------------
# Paragraph text reconstruction (preserve math/equations)
# ------------------------------
def paragraph_full_text(paragraph) -> str:
    """
    Return paragraph text, concatenating w:t and m:t nodes in order.
    Preserves equations as linear text.
    """
    xml = paragraph._p.xml.encode("utf-8")
    root = etree.fromstring(xml)
    parts = []
    for node in root:
        tag = etree.QName(node.tag).localname
        if tag == "r":
            for t in node.findall(".//w:t", namespaces=NSMAP):
                if t.text:
                    parts.append(t.text)
            if node.findall(".//w:br", namespaces=NSMAP):
                parts.append("\n")
        elif tag in ("oMath", "oMathPara"):
            for mt in node.findall(".//m:t", namespaces=NSMAP):
                if mt.text:
                    parts.append(mt.text)
        else:
            for t in node.findall(".//w:t", namespaces=NSMAP):
                if t.text:
                    parts.append(t.text)
    return "".join(parts).strip()

# ------------------------------
# Normalize text
# ------------------------------
def normalize_text(s: Optional[str]) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFC", s)
    s = s.replace("\u200b", "").replace("\u200c", "").replace("\u200d", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# ------------------------------
# Block grouping
# ------------------------------
SEPARATOR_RE = re.compile(r"^[\-\—\–\*\_]{1,}\s*$")

def group_blocks(paragraphs: List[str]) -> List[List[str]]:
    blocks, current = [], []
    for p in paragraphs:
        if not p or SEPARATOR_RE.match(p.strip()):
            if current:
                blocks.append(current)
                current = []
            continue
        current.append(p)
    if current:
        blocks.append(current)
    return blocks

# ------------------------------
# Parse block into question dict
# ------------------------------
def parse_block(block: List[str], idx: int) -> Optional[dict]:
    text = " ".join(block).strip()
    if not text:
        return None

    # Remove heading-like blocks
    if re.match(r"(?i)^(unique questions|paper \d+|selected questions)", text):
        return None

    debug = {"block_index": idx, "raw": text}

    # Extract explanation
    explanation = ""
    m_expl = re.search(r"(?i)(Explanation|Solution)\s*[:\-]?\s*(.*)$", text)
    if m_expl:
        explanation = m_expl.group(2).strip()
        text = text[:m_expl.start()].strip()
    debug["explanation"] = explanation

    # Extract answer
    answer_letter, raw_answer_text = None, None
    m_ans = re.search(r"(?i)(Answer|Ans|Correct|Key)\s*[:\-]?\s*([^\n\r]*)", text)
    if m_ans:
        raw_answer_text = m_ans.group(2).strip()
        text = text[:m_ans.start()].strip()
        m_letter = re.search(r"([A-Da-d])", raw_answer_text)
        if m_letter:
            answer_letter = m_letter.group(1).lower()
    debug["raw_answer"] = raw_answer_text

    # Split question vs options
    qtext, options = text, []
    m_opt = re.search(r"(?i)\bOptions?\b\s*[:\-]?\s*(.*)$", text)
    if m_opt:
        qtext = text[:m_opt.start()].strip()
        options_part = m_opt.group(1).strip()
    else:
        sp = re.split(r"(?=(?:\(|\[)?[A-Da-d][\)\].]?\s+)", text, maxsplit=1)
        if len(sp) == 2:
            qtext, options_part = sp[0].strip(), sp[1].strip()
        else:
            options_part = ""

    # Extract options
    pairs = re.findall(
        r"[\(\[]?([A-Da-d])[\)\].]?\s*"
        r"([^(\(\[]+?)"
        r"(?=(?:[\(\[]?[A-Da-d][\)\].]?|\Z))",
        options_part, flags=re.S
    )
    options = [normalize_text(opt) for _, opt in pairs]
    while len(options) < 4:
        options.append("")
    options = options[:4]
    debug["options"] = options

    # Resolve answer
    assumed = False
    if not answer_letter and raw_answer_text:
        for i, opt in enumerate(options):
            if opt and opt.lower() in raw_answer_text.lower():
                answer_letter = "abcd"[i]
                break
    if not answer_letter:
        answer_letter, assumed = "a", True
    debug["answer"] = answer_letter
    debug["assumed"] = assumed

    return {
        "question": normalize_text(qtext),
        "options": options,
        "answer": answer_letter,
        "assumed": assumed,
        "explanation": normalize_text(explanation),
        "debug": debug,
    }

# ------------------------------
# Parse docx into questions
# ------------------------------
def parse_docx(input_path: str) -> List[dict]:
    doc = Document(input_path)
    paras = [paragraph_full_text(p) for p in doc.paragraphs]
    blocks = group_blocks(paras)
    questions = []
    with open("debug_log.jsonl", "w", encoding="utf-8") as dbg:
        for i, block in enumerate(blocks):
            q = parse_block(block, i)
            if q:
                questions.append(q)
                dbg.write(json.dumps(q["debug"], ensure_ascii=False) + "\n")
    return questions

# ------------------------------
# Output formatter
# ------------------------------
def write_output(questions: List[dict], output_path: str):
    doc = Document()
    for q in questions:
        table = doc.add_table(rows=8, cols=3)
        try:
            table.style = "Table Grid"
        except Exception:
            pass

        # Row 0: Question
        table.cell(0, 0).text = "Question"
        table.cell(0, 1).merge(table.cell(0, 2)).text = q["question"]

        # Row 1: Type
        table.cell(1, 0).text = "Type"
        table.cell(1, 1).merge(table.cell(1, 2)).text = "multiple_choice"

        # Rows 2-5: Options
        answer_idx = {"a": 0, "b": 1, "c": 2, "d": 3}.get(q["answer"], 0)
        for i in range(4):
            table.cell(2 + i, 0).text = "Option"
            table.cell(2 + i, 1).text = q["options"][i]
            table.cell(2 + i, 2).text = "correct" if i == answer_idx else "incorrect"

        # Row 6: Solution
        table.cell(6, 0).text = "Solution"
        table.cell(6, 1).merge(table.cell(6, 2)).text = q["explanation"]

        # Row 7: Marks
        table.cell(7, 0).text = "Marks"
        table.cell(7, 1).text = "1"
        table.cell(7, 2).text = "0"

        doc.add_paragraph()
    doc.save(output_path)

# ------------------------------
# Main
# ------------------------------
def main():
    if len(sys.argv) < 3:
        print("Usage: python quizformatter_fixed.py input.docx output.docx")
        return
    input_path, output_path = sys.argv[1], sys.argv[2]
    questions = parse_docx(input_path)
    print(f"Parsed {len(questions)} questions.")
    write_output(questions, output_path)
    print(f"Saved to {output_path}. Debug log: debug_log.jsonl")

if __name__ == "__main__":
    main()
