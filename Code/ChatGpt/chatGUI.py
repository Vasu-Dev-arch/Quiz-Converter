# Quiz Converter - Desktop UI (PySide6)
# Replicates the clean website mockup we designed earlier.
# Features:
# - Top bar with app name (top-left) and settings (top-right)
# - Centered card with Input and Output fields
# - Labels above inputs, thin bottom border style for inputs
# - Pill-shaped Browse / Save / Convert buttons (Browse & Save equal size)
# - Theme toggle (Light / Dark) with distinct contrast
# - File dialogs for selecting input file and choosing save location
# - Simple conversion logic that attempts to parse MCQ-like text into JSON
# - Small toast message for feedback
# - Keyboard shortcut Ctrl+K to focus output name

# Requirements:
# pip install PySide6

import sys
import json
import re
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QAction, QIcon, QKeySequence, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox, QFrame, QSizePolicy
)


LIGHT_STYLESHEET = """
/* Light theme */
QWidget { background: #f5f5f5; color: #1f2933; }
#topbar { background: #222; color: #fff; }
#card { background: #ffffff; border-radius: 14px; }
QLabel { color: #334155; }
QLineEdit { background: transparent; border: none; color: #0b1220; font-size: 16px; }
QPushButton.pill { background: #007bff; color: white; border-radius: 20px; padding: 8px 18px; }
QPushButton.pill:hover { background: #0056b3; }
QPushButton.convert { background: #007bff; }
#input_line { border-bottom: 1px solid #cfcfcf; padding-bottom: 8px; }
#output_line { border-bottom: 1px solid #cfcfcf; padding-bottom: 8px; }
#app_name { color: white; font-weight: 700; }
"""

DARK_STYLESHEET = """
/* Dark theme */
QWidget { background: #121212; color: #e6eef8; }
#topbar { background: #111827; color: #fff; }
#card { background: #1f2937; border-radius: 14px; }
QLabel { color: #d1d5db; }
QLineEdit { background: transparent; border: none; color: #fff; font-size: 16px; }
QPushButton.pill { background: #0a84ff; color: white; border-radius: 20px; padding: 8px 18px; }
QPushButton.pill:hover { background: #0066cc; }
QPushButton.convert { background: #0a84ff; }
#input_line { border-bottom: 1px solid #374151; padding-bottom: 8px; }
#output_line { border-bottom: 1px solid #374151; padding-bottom: 8px; }
#app_name { color: white; font-weight: 800; }
"""


class QuizConverterMain(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quiz Converter")
        self.setMinimumSize(680, 420)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        # central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 16)
        main_layout.setSpacing(12)

        # top bar
        topbar = QWidget(objectName='topbar')
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(18, 12, 18, 12)

        self.app_name = QLabel("Quiz Converter", objectName='app_name')
        self.app_name.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        topbar_layout.addWidget(self.app_name, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        topbar_layout.addStretch()

        # settings button
        self.settings_btn = QPushButton("⚙️")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setFixedSize(36, 36)
        self.settings_btn.setToolTip("Toggle theme")
        self.settings_btn.clicked.connect(self.toggle_theme)
        topbar_layout.addWidget(self.settings_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)

        main_layout.addWidget(topbar)

        # content area: center the card
        content = QWidget()
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        content_layout.addStretch()

        card = QFrame(objectName='card')
        card.setFixedWidth(420)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(36, 28, 36, 28)
        card_layout.setSpacing(14)

        # Input group
        lbl_input = QLabel("Input")
        lbl_input.setFont(QFont("Arial", 13))
        card_layout.addWidget(lbl_input)

        input_line = QWidget(objectName='input_line')
        input_line_layout = QHBoxLayout(input_line)
        input_line_layout.setContentsMargins(6, 0, 6, 6)
        input_line_layout.setSpacing(8)

        self.input_edit = QLineEdit(placeholderText='Choose file...')
        self.input_edit.setReadOnly(True)
        self.input_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self.browse_btn = QPushButton("Browse")
        self.browse_btn.setObjectName('browse_btn')
        self.browse_btn.setProperty('class', 'pill')
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.clicked.connect(self.browse_file)
        self.browse_btn.setFixedWidth(110)

        input_line_layout.addWidget(self.input_edit)
        input_line_layout.addWidget(self.browse_btn)
        card_layout.addWidget(input_line)

        # Output group
        lbl_output = QLabel("Output")
        lbl_output.setFont(QFont("Arial", 13))
        card_layout.addWidget(lbl_output)

        output_line = QWidget(objectName='output_line')
        output_line_layout = QHBoxLayout(output_line)
        output_line_layout.setContentsMargins(6, 0, 6, 6)
        output_line_layout.setSpacing(8)

        self.output_edit = QLineEdit(placeholderText='converted-quiz.json')
        self.output_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        self.save_btn = QPushButton("Save")
        self.save_btn.setObjectName('save_btn')
        self.save_btn.setProperty('class', 'pill')
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self.save_as)
        self.save_btn.setFixedWidth(110)

        output_line_layout.addWidget(self.output_edit)
        output_line_layout.addWidget(self.save_btn)
        card_layout.addWidget(output_line)

        # Convert button
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setProperty('class', 'pill')
        self.convert_btn.setObjectName('convert_btn')
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.setFixedHeight(44)
        self.convert_btn.clicked.connect(self.convert)
        self.convert_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        card_layout.addSpacing(6)
        card_layout.addWidget(self.convert_btn)

        # stretch to keep compact
        card_layout.addStretch()

        content_layout.addWidget(card, alignment=Qt.AlignCenter)
        content_layout.addStretch()

        main_layout.addWidget(content)

        # toast label
        self.toast = QLabel("")
        self.toast.setStyleSheet("padding:10px; border-radius:10px; background:#007bff; color:white;")
        self.toast.setVisible(False)
        self.toast.setFixedWidth(300)
        main_layout.addWidget(self.toast, alignment=Qt.AlignRight | Qt.AlignBottom)

        # default theme
        self.is_dark = False
        self.apply_theme()

        # shortcut
        focus_output = QAction(self)
        focus_output.setShortcut(QKeySequence("Ctrl+K"))
        focus_output.triggered.connect(lambda: self.output_edit.setFocus())
        self.addAction(focus_output)

        # internal state
        self.current_input_path = None

    def show_toast(self, text, ms=2200):
        self.toast.setText(text)
        self.toast.setVisible(True)
        QTimer.singleShot(ms, lambda: self.toast.setVisible(False))

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.apply_theme()

    def apply_theme(self):
        if self.is_dark:
            self.setStyleSheet(DARK_STYLESHEET)
        else:
            self.setStyleSheet(LIGHT_STYLESHEET)
        # tweak app-name font size for clarity
        self.app_name.setFont(QFont("Arial", 22, QFont.Weight.DemiBold))

    def browse_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select input file", str(Path.home()), "Text Files (*.txt *.md);;CSV Files (*.csv);;JSON Files (*.json);;All Files (*)")
        if fname:
            self.current_input_path = Path(fname)
            self.input_edit.setText(self.current_input_path.name)
            # suggest output name if empty
            if not self.output_edit.text().strip():
                suggested = self.current_input_path.stem + "-converted.json"
                self.output_edit.setText(suggested)
            self.show_toast("File loaded: " + self.current_input_path.name, 1600)

    def save_as(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Save output as", str(Path.home() / self.output_edit.text()), "JSON Files (*.json);;All Files (*)")
        if fname:
            self.output_edit.setText(Path(fname).name)
            self.output_path = Path(fname)
            self.show_toast("Save location set", 1400)

    def convert(self):
        if not self.current_input_path or not self.current_input_path.exists():
            QMessageBox.warning(self, "No input", "Please choose a valid input file first.")
            return

        out_name = self.output_edit.text().strip() or (self.current_input_path.stem + "-converted.json")
        # If user previously used Save As, prefer that path
        out_path = getattr(self, 'output_path', None)
        if not out_path:
            # default to same folder as input
            out_path = self.current_input_path.parent / out_name

        try:
            text = self.current_input_path.read_text(encoding='utf-8', errors='ignore')
        except Exception as e:
            QMessageBox.critical(self, "Read error", f"Failed to read input file:\n{e}")
            return

        parsed = self.simple_convert(text)
        payload = {
            'sourceFile': str(self.current_input_path.name),
            'convertedAt': datetime.utcnow().isoformat() + 'Z',
            'converter': 'Quiz Converter (Desktop)',
            'summary': parsed.get('summary', ''),
            'data': parsed.get('data')
        }

        try:
            out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding='utf-8')
            self.show_toast(f"Converted and saved: {out_path.name}")
        except Exception as e:
            QMessageBox.critical(self, "Save error", f"Failed to save output file:\n{e}")

    def simple_convert(self, text: str):
        # Logic inspired by the JS simpleConvert in the prototype
        lines = [l.strip() for l in text.splitlines()]
        lines = [l for l in lines if l]
        questions = []
        i = 0
        while i < len(lines):
            l = lines[i]
            qMatch = re.match(r'^(?:\d+\.|Q:|Question\s*[:\-]?)(.*)$', l, re.IGNORECASE)
            if qMatch:
                qtext = qMatch.group(1).strip()
                i += 1
                options = []
                answer = None
                while i < len(lines) and not re.match(r'^(?:\d+\.|Q:|Question\s*[:\-]?)', lines[i], re.IGNORECASE):
                    opt = lines[i]
                    a = re.match(r'^(?:A\.|B\.|C\.|D\.|[A-D]\)|\-|•)\s*(.*)$', opt, re.IGNORECASE)
                    ans = re.match(r'^(?:Answer[:\-]?|Ans[:\-]?|Correct[:\-]?)(.*)$', opt, re.IGNORECASE)
                    if a:
                        options.append(a.group(1).strip())
                    elif ans:
                        answer = ans.group(1).strip()
                    else:
                        if options:
                            options[-1] += ' ' + opt
                        else:
                            qtext += ' ' + opt
                    i += 1
                questions.append({'question': qtext, 'options': options, 'answer': answer})
            else:
                i += 1

        if questions:
            return {'summary': f'Parsed {len(questions)} question(s)', 'data': questions}
        else:
            return {'summary': 'No structured questions found — packaged raw text', 'data': {'raw': text}}


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Quiz Converter")

    window = QuizConverterMain()
    window.show()

    # center the window on screen
    screen = app.primaryScreen().availableGeometry()
    size = window.geometry()
    window.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 4)

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
