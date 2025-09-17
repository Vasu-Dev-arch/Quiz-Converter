# quiz_converter_app_fixed.py
# PySide6 desktop UI — fixed styles and layout to exactly match the web mock.
# pip install PySide6
import sys
import json
import re
from pathlib import Path
from datetime import datetime
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QKeySequence, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox, QFrame, QSizePolicy
)

# -----------------------
# Application-level QSS
# -----------------------
LIGHT_QSS = r"""
/* Base + font */
QWidget { background: #f5f5f5; color: #0f1720; font-family: "Segoe UI", Arial, sans-serif; }

/* Top bar */
#topbar { background: #111827; min-height: 64px; max-height: 64px; }
#app_name { background: transparent; color: #ffffff; font-weight: 700; font-size: 30px; padding-left:6px; }

/* Card */
#card { background: #ffffff; border-radius: 12px; }

/* Labels & inputs */
QLabel { background: transparent; color: #374151; font-size: 15px; }
QLineEdit { background: transparent; border: none; color: #0f1720; font-size: 15px; padding: 6px 4px; }

/* Thin bottom border containers */
#input_line, #output_line {
  border-bottom: 1px solid #e6e6e6;
  padding-bottom: 8px;
}

/* Pill buttons using string dynamic property pclass="pill" */
QPushButton[pclass="pill"] {
  background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #007bff, stop:1 #007bff);
  color: white;
  border-radius: 999px;          /* perfect smooth pill */
  min-width: 110px;
  min-height: 36px;
  font-size: 15px;
  padding: 6px 14px;
  border: none;
}
QPushButton[pclass="pill"]:hover { background: #0056b3; }

/* Convert button specifics (object name convert_btn) - slightly reduced height */
QPushButton#convert_btn[pclass="pill"] {
  min-height: 40px;            /* reduced a bit */
  font-size: 15px;
  padding: 8px 16px;
}

/* Settings button */
QPushButton#settings_btn {
  background: transparent;
  color: #e5e7eb;
  border: none;
  min-width: 36px;
  min-height: 36px;
  border-radius: 8px;
}

/* Toast */
#toast {
  background: rgba(0,0,0,0.75);
  color: #fff;
  padding: 10px 14px;
  border-radius: 10px;
}

/* Remove frames */
QFrame { border: none; }
"""

DARK_QSS = r"""
/* Base + font */
QWidget { background: #0b0f13; color: #e6eef8; font-family: "Segoe UI", Arial, sans-serif; }

/* Top bar */
#topbar { background: #111827; min-height: 64px; max-height: 64px; }
#app_name { background: transparent; color: #fff; font-weight: 700; font-size: 30px; padding-left:6px; }

/* Card */
#card { background: #111827; border-radius: 12px; }

/* Labels & inputs */
QLabel { background: transparent; color: #d1d5db; font-size: 15px; }
QLineEdit { background: transparent; border: none; color: #fff; font-size: 15px; padding: 6px 4px; }

/* Thin bottom border containers */
#input_line, #output_line {
  border-bottom: 1px solid #374151;
  padding-bottom: 8px;
}

/* Pill buttons */
QPushButton[pclass="pill"] {
  background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #0a84ff, stop:1 #0a84ff);
  color: white;
  border-radius: 999px;
  min-width: 110px;
  min-height: 36px;
  font-size: 15px;
  padding: 6px 14px;
  border: none;
}
QPushButton[pclass="pill"]:hover { background: #0066cc; }

/* Convert button specifics */
QPushButton#convert_btn[pclass="pill"] {
  min-height: 40px;
  font-size: 15px;
  padding: 8px 16px;
}

/* Settings button */
QPushButton#settings_btn {
  background: rgba(255,255,255,0.04);
  color: #d1d5db;
  border: none;
  min-width: 36px;
  min-height: 36px;
  border-radius: 8px;
}

/* Toast */
#toast {
  background: rgba(255,255,255,0.08);
  color: #e6eef8;
  padding: 10px 14px;
  border-radius: 10px;
}

QFrame { border: none; }
"""

# -----------------------
# Main window
# -----------------------
class QuizConverterMain(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quiz Converter")
        self.setMinimumSize(820, 520)

        # central widget and layout
        central = QWidget()
        self.setCentralWidget(central)
        main_v = QVBoxLayout(central)
        main_v.setContentsMargins(0, 0, 0, 12)
        main_v.setSpacing(0)

        # -------- Top bar (compact) --------
        topbar = QWidget(objectName="topbar")
        topbar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(18, 10, 18, 10)
        topbar_layout.setSpacing(8)

        # App name (font-size controlled via QSS #app_name)
        self.app_name = QLabel("Quiz Converter", objectName="app_name")
        # still set a fallback font (QSS will control size but keep weight)
        self.app_name.setFont(QFont("Segoe UI", 26, QFont.Weight.Bold))
        topbar_layout.addWidget(self.app_name, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        topbar_layout.addStretch()

        # settings button (top-right)
        self.settings_btn = QPushButton("⚙", objectName="settings_btn")
        self.settings_btn.setToolTip("Toggle theme")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setFixedSize(36, 36)
        self.settings_btn.clicked.connect(self.toggle_theme)
        topbar_layout.addWidget(self.settings_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)

        main_v.addWidget(topbar)

        # -------- Content area (centered card) --------
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 18, 0, 18)
        content_layout.setSpacing(0)

        content_layout.addStretch()  # top spacer to help vertical centering

        # horizontal wrapper to center the card
        hwrap = QHBoxLayout()
        hwrap.addStretch()

        card = QFrame(objectName="card")
        card.setFixedWidth(420)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(28, 22, 28, 22)
        card_layout.setSpacing(14)

        # Input label + input row
        lbl_input = QLabel("Input")
        lbl_input.setFont(QFont("Segoe UI", 13))
        card_layout.addWidget(lbl_input)

        input_line = QWidget(objectName="input_line")
        input_layout = QHBoxLayout(input_line)
        input_layout.setContentsMargins(6, 0, 6, 6)
        input_layout.setSpacing(12)  # ensure button is beside field

        self.input_edit = QLineEdit(placeholderText="Choose file...")
        self.input_edit.setReadOnly(True)
        self.input_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.input_edit.setFont(QFont("Segoe UI", 13))

        self.browse_btn = QPushButton("Browse")
        # use a string dynamic property (more robust than bool)
        self.browse_btn.setProperty("pclass", "pill")
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.setFixedWidth(110)
        self.browse_btn.setFixedHeight(36)
        self.browse_btn.clicked.connect(self.browse_file)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(self.browse_btn)
        card_layout.addWidget(input_line)

        # Output label + row
        lbl_output = QLabel("Output")
        lbl_output.setFont(QFont("Segoe UI", 13))
        card_layout.addWidget(lbl_output)

        output_line = QWidget(objectName="output_line")
        output_layout = QHBoxLayout(output_line)
        output_layout.setContentsMargins(6, 0, 6, 6)
        output_layout.setSpacing(12)

        self.output_edit = QLineEdit(placeholderText="Save as...")
        self.output_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.output_edit.setFont(QFont("Segoe UI", 13))

        self.save_btn = QPushButton("Save")
        self.save_btn.setProperty("pclass", "pill")
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.setFixedWidth(110)
        self.save_btn.setFixedHeight(36)
        self.save_btn.clicked.connect(self.save_as)

        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(self.save_btn)
        card_layout.addWidget(output_line)

        # Convert button (full width pill, slightly reduced height)
        self.convert_btn = QPushButton("Convert", objectName="convert_btn")
        self.convert_btn.setProperty("pclass", "pill")
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.setFixedHeight(40)  # reduced a little
        self.convert_btn.clicked.connect(self.convert)
        card_layout.addSpacing(6)
        card_layout.addWidget(self.convert_btn)

        card_layout.addStretch()
        hwrap.addWidget(card)
        hwrap.addStretch()
        content_layout.addLayout(hwrap)

        content_layout.addStretch()  # bottom spacer for vertical centering
        main_v.addWidget(content)

        # Toast (bottom center)
        self.toast = QLabel("", objectName="toast")
        self.toast.setVisible(False)
        self.toast.setAlignment(Qt.AlignCenter)
        toast_wrap = QHBoxLayout()
        toast_wrap.addStretch()
        toast_wrap.addWidget(self.toast)
        toast_wrap.addStretch()
        main_v.addLayout(toast_wrap)

        # state
        self.is_dark = False
        self.current_input_path = None
        self.output_path = None

        # apply theme at application level (robust)
        self.apply_theme()

        # keyboard shortcut Ctrl+K to focus output field
        focus_output = QAction(self)
        focus_output.setShortcut(QKeySequence("Ctrl+K"))
        focus_output.triggered.connect(lambda: self.output_edit.setFocus())
        self.addAction(focus_output)

    # ----------------------
    # Theming & helpers
    # ----------------------
    def apply_theme(self):
        """Apply stylesheet at application-level so it can't be easily overridden by platform style."""
        app = QApplication.instance()
        if app is None:
            return
        if self.is_dark:
            app.setStyleSheet(DARK_QSS)
        else:
            app.setStyleSheet(LIGHT_QSS)
        # also ensure the label font weight remains bold through QSS
        self.app_name.setFont(QFont("Segoe UI", 26, QFont.Weight.Bold))

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.apply_theme()

    def show_toast(self, text: str, ms: int = 2000):
        self.toast.setText(text)
        self.toast.setVisible(True)
        QTimer.singleShot(ms, lambda: self.toast.setVisible(False))

    # ----------------------
    # File actions
    # ----------------------
    def browse_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select input file", str(Path.home()),
                                              "Text Files (*.txt *.md);;Word Documents (*.docx);;All Files (*)")
        if not fname:
            return
        self.current_input_path = Path(fname)
        self.input_edit.setText(self.current_input_path.name)
        # suggest a filename but don't force .json in the placeholder
        if not self.output_edit.text().strip():
            self.output_edit.setText(self.current_input_path.stem + "-converted")
        self.show_toast(f"Loaded: {self.current_input_path.name}", 1400)

    def save_as(self):
        default_name = (self.output_edit.text().strip() or
                        (self.current_input_path.stem + "-converted" if self.current_input_path else "converted-quiz"))
        fname, _ = QFileDialog.getSaveFileName(self, "Save output as", str(Path.home() / default_name),
                                              "JSON Files (*.json);;All Files (*)")
        if not fname:
            return
        self.output_path = Path(fname)
        # show only filename in UI (matches web mock behavior)
        self.output_edit.setText(self.output_path.name)
        self.show_toast("Save location set", 1200)

    # ----------------------
    # Convert (plug backend here)
    # ----------------------
    def convert(self):
        if not self.current_input_path or not self.current_input_path.exists():
            QMessageBox.warning(self, "No input", "Please choose a valid input file first.")
            return

        # compute output path
        out_name_text = self.output_edit.text().strip()
        if self.output_path:
            out_path = self.output_path
        else:
            out_path = self.current_input_path.parent / (out_name_text or (self.current_input_path.stem + "-converted.json"))

        try:
            text = self.current_input_path.read_text(encoding="utf-8", errors="ignore")
        except Exception as ex:
            QMessageBox.critical(self, "Read error", f"Failed to read input file:\n{ex}")
            return

        # === PLUG YOUR BACKEND PROCESS HERE ===
        # Example placeholder: we run a simple parser (replace with your word-to-word converter)
        parsed = self.simple_convert(text)
        payload = {
            "sourceFile": str(self.current_input_path.name),
            "convertedAt": datetime.utcnow().isoformat() + "Z",
            "converter": "Quiz Converter (Desktop)",
            "summary": parsed.get("summary", ""),
            "data": parsed.get("data")
        }
        # =======================================

        try:
            out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            self.show_toast(f"Converted and saved: {out_path.name}", 1800)
        except Exception as ex:
            QMessageBox.critical(self, "Save error", f"Failed to save output file:\n{ex}")

    def simple_convert(self, text: str):
        # same simpleMCQ parser as placeholder
        lines = [l.strip() for l in text.splitlines() if l.strip()]
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
                questions.append({"question": qtext, "options": options, "answer": answer})
            else:
                i += 1
        if questions:
            return {"summary": f"Parsed {len(questions)} question(s)", "data": questions}
        return {"summary": "No structured questions found — packaged raw text", "data": {"raw": text}}


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Quiz Converter")

    # instantiate window after app so we can call app.setStyleSheet() in apply_theme
    window = QuizConverterMain()
    window.show()

    # center window on screen
    screen = app.primaryScreen().availableGeometry()
    win_geom = window.geometry()
    window.move((screen.width() - win_geom.width()) // 2, (screen.height() - win_geom.height()) // 2)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
