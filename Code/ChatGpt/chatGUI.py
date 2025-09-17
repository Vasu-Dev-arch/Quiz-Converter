# quiz_converter_app_windows_ready.py
# PySide6 desktop UI — Windows-ready, dark-theme default, true pill buttons, aligned inputs/buttons.
# pip install PySide6

import sys, json, re
from pathlib import Path
from datetime import datetime
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QKeySequence, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox, QFrame, QSizePolicy
)

# ------------ QSS (light & dark) ------------
# We keep both in case you want to switch default later.
DARK_QSS = r"""
/* Base background and font */
QWidget { background: #0b0f13; color: #e6eef8; font-family: "Segoe UI", Arial, sans-serif; }

/* Topbar */
#topbar { background: #0f1720; min-height: 64px; max-height: 64px; }
#app_name { color: #ffffff; font-weight: 800; font-size: 32px; padding-left:6px; }

/* Card */
#card { background: #0f1728; border-radius: 12px; }

/* Labels & inputs */
QLabel { background: transparent; color: #cbd5e1; font-size: 15px; }
QLineEdit { background: transparent; border: none; color: #e6eef8; font-size: 15px; padding: 6px 8px; border-bottom: 1px solid #1f2937; }

/* Thin bottom border containers to mimic modern underline input */
#input_line, #output_line {
  padding-bottom: 8px;
}

/* Pill buttons: using dynamic property pclass="pill" */
QPushButton[pclass="pill"] {
  background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #0a84ff, stop:1 #0a84ff);
  color: white;
  border-radius: 999px;
  min-width: 110px;
  min-height: 36px;
  font-size: 16px;
  padding: 6px 16px;
  border: none;
}
QPushButton[pclass="pill"]:hover { background: #0066cc; }

/* Convert button slightly smaller height than before for your request */
QPushButton#convert_btn[pclass="pill"] {
  min-height: 40px;
  font-size: 16px;
  padding: 8px 18px;
}

/* Theme toggle button */
QPushButton#theme_btn {
  background: rgba(255,255,255,0.04);
  color: #d1d5db;
  border: none;
  min-width: 40px;
  min-height: 40px;
  border-radius: 8px;
}

/* Toast */
#toast {
  background: rgba(255,255,255,0.06);
  color: #e6eef8;
  padding: 10px 14px;
  border-radius: 10px;
}

/* Remove frame borders globally (we use QFrame as container) */
QFrame { border: none; }
"""

LIGHT_QSS = r"""
QWidget { background: #f5f5f5; color: #0f1720; font-family: "Segoe UI", Arial, sans-serif; }

/* Topbar */
#topbar { background: #111827; min-height: 64px; max-height: 64px; }
#app_name { color: #ffffff; font-weight: 800; font-size: 32px; padding-left:6px; }

/* Card */
#card { background: #ffffff; border-radius: 12px; }

/* Labels & inputs */
QLabel { background: transparent; color: #374151; font-size: 15px; }
QLineEdit { background: transparent; border: none; color: #0f1720; font-size: 15px; padding: 6px 8px; border-bottom: 1px solid #e6e6e6; }

/* underline style */
#input_line, #output_line {
  padding-bottom: 8px;
}

/* pill */
QPushButton[pclass="pill"] {
  background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #007bff, stop:1 #007bff);
  color: white;
  border-radius: 999px;
  min-width: 110px;
  min-height: 36px;
  font-size: 16px;
  padding: 6px 16px;
  border: none;
}
QPushButton[pclass="pill"]:hover { background: #0056b3; }

QPushButton#convert_btn[pclass="pill"] {
  min-height: 40px;
  font-size: 16px;
  padding: 8px 18px;
}

/* Theme toggle */
QPushButton#theme_btn {
  background: rgba(0,0,0,0.06);
  color: #111827;
  border: none;
  min-width: 40px;
  min-height: 40px;
  border-radius: 8px;
}

/* Toast */
#toast { background: rgba(0,0,0,0.65); color: #fff; padding: 10px 14px; border-radius: 10px; }
QFrame { border: none; }
"""

# -------------------- Main Window --------------------
class QuizConverterMain(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quiz Converter")
        self.setMinimumSize(880, 560)

        # central layout
        central = QWidget()
        self.setCentralWidget(central)
        main_v = QVBoxLayout(central)
        main_v.setContentsMargins(0, 0, 0, 14)
        main_v.setSpacing(0)

        # ----- Topbar (compact) -----
        topbar = QWidget(objectName="topbar")
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(18, 10, 18, 10)
        topbar_layout.setSpacing(8)

        self.app_name = QLabel("Quiz Converter", objectName="app_name")
        # fallback font; QSS will set size, but keep weight
        self.app_name.setFont(QFont("Segoe UI", 28, QFont.Weight.Bold))
        topbar_layout.addWidget(self.app_name, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        topbar_layout.addStretch()

        # Theme toggle button (shows moon or sun) — more intuitive than gear
        self.theme_btn = QPushButton("🌙", objectName="theme_btn")
        self.theme_btn.setCursor(Qt.PointingHandCursor)
        self.theme_btn.setFixedSize(40, 40)
        self.theme_btn.clicked.connect(self.toggle_theme)
        # will set pclass not needed for theme icon
        topbar_layout.addWidget(self.theme_btn, alignment=Qt.AlignRight | Qt.AlignVCenter)

        main_v.addWidget(topbar)

        # ----- Content area — centered card -----
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 18, 0, 18)
        content_layout.setSpacing(0)

        content_layout.addStretch()

        # center horizontally
        hwrap = QHBoxLayout()
        hwrap.addStretch()

        card = QFrame(objectName="card")
        card.setFixedWidth(480)  # slightly wider for comfortable spacing
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(32, 26, 32, 26)
        card_layout.setSpacing(16)

        # Input
        lbl_input = QLabel("Input")
        lbl_input.setFont(QFont("Segoe UI", 13))
        card_layout.addWidget(lbl_input)

        input_line = QWidget(objectName="input_line")
        input_layout = QHBoxLayout(input_line)
        input_layout.setContentsMargins(6, 0, 6, 6)
        input_layout.setSpacing(12)

        self.input_edit = QLineEdit(placeholderText="Choose file...")
        # match button height to align perfectly
        self.input_edit.setFixedHeight(36)
        self.input_edit.setReadOnly(True)
        self.input_edit.setFont(QFont("Segoe UI", 13))
        self.input_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.browse_btn = QPushButton("Browse")
        self.browse_btn.setProperty("pclass", "pill")
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.setFixedHeight(36)
        self.browse_btn.setFixedWidth(120)
        # ensure native frames don't apply
        self.browse_btn.setFlat(True)
        self.browse_btn.clicked.connect(self.browse_file)

        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(self.browse_btn)
        card_layout.addWidget(input_line)

        # Output
        lbl_output = QLabel("Output")
        lbl_output.setFont(QFont("Segoe UI", 13))
        card_layout.addWidget(lbl_output)

        output_line = QWidget(objectName="output_line")
        output_layout = QHBoxLayout(output_line)
        output_layout.setContentsMargins(6, 0, 6, 6)
        output_layout.setSpacing(12)

        self.output_edit = QLineEdit(placeholderText="Save as...")
        self.output_edit.setFixedHeight(36)
        self.output_edit.setFont(QFont("Segoe UI", 13))
        self.output_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.save_btn = QPushButton("Save")
        self.save_btn.setProperty("pclass", "pill")
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.setFixedHeight(36)
        self.save_btn.setFixedWidth(120)
        self.save_btn.setFlat(True)
        self.save_btn.clicked.connect(self.save_as)

        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(self.save_btn)
        card_layout.addWidget(output_line)

        # Convert button — full width, slightly reduced height, pill
        self.convert_btn = QPushButton("Convert", objectName="convert_btn")
        self.convert_btn.setProperty("pclass", "pill")
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.setFixedHeight(40)
        self.convert_btn.setFlat(True)
        self.convert_btn.clicked.connect(self.convert)
        card_layout.addSpacing(6)
        card_layout.addWidget(self.convert_btn)

        card_layout.addStretch()
        hwrap.addWidget(card)
        hwrap.addStretch()
        content_layout.addLayout(hwrap)

        content_layout.addStretch()
        main_v.addWidget(content)

        # Toast
        self.toast = QLabel("", objectName="toast")
        self.toast.setVisible(False)
        self.toast.setAlignment(Qt.AlignCenter)
        toast_wrap = QHBoxLayout()
        toast_wrap.addStretch()
        toast_wrap.addWidget(self.toast)
        toast_wrap.addStretch()
        main_v.addLayout(toast_wrap)

        # state
        self.current_input_path = None
        self.output_path = None

        # default to dark theme as requested
        self.is_dark = True

        # IMPORTANT: use Fusion style so QSS reliably applies across Windows versions
        app = QApplication.instance()
        app.setStyle("Fusion")
        # apply theme (app-level stylesheet) now
        self.apply_theme()

        # shortcuts
        focus_output = QAction(self)
        focus_output.setShortcut(QKeySequence("Ctrl+K"))
        focus_output.triggered.connect(lambda: self.output_edit.setFocus())
        self.addAction(focus_output)

        # ensure theme button icon matches initial theme
        self.update_theme_icon()

    # ---------- Styling helpers ----------
    def apply_theme(self):
        app = QApplication.instance()
        if app is None:
            return
        if self.is_dark:
            app.setStyleSheet(DARK_QSS)
        else:
            app.setStyleSheet(LIGHT_QSS)
        # ensure app-name weight & font stays strong
        self.app_name.setFont(QFont("Segoe UI", 30, QFont.Weight.DemiBold))
        self.update_theme_icon()

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.apply_theme()

    def update_theme_icon(self):
        # show moon for dark, sun for light
        if self.is_dark:
            self.theme_btn.setText("☀")  # shows sun icon to indicate "switch to light"
            self.theme_btn.setToolTip("Switch to light theme")
        else:
            self.theme_btn.setText("🌙")  # shows moon icon to indicate "switch to dark"
            self.theme_btn.setToolTip("Switch to dark theme")

    def show_toast(self, text: str, ms: int = 2000):
        self.toast.setText(text)
        self.toast.setVisible(True)
        QTimer.singleShot(ms, lambda: self.toast.setVisible(False))

    # ---------- File actions ----------
    def browse_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select input file", str(Path.home()),
                                              "Text Files (*.txt *.md);;Word Documents (*.docx);;All Files (*)")
        if not fname:
            return
        self.current_input_path = Path(fname)
        self.input_edit.setText(self.current_input_path.name)
        # suggest output (no forced .json placeholder)
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
        # show filename only, like the web mock
        self.output_edit.setText(self.output_path.name)
        self.show_toast("Save location set", 1200)

    # ---------- Convert (plug backend here) ----------
    def convert(self):
        if not self.current_input_path or not self.current_input_path.exists():
            QMessageBox.warning(self, "No input", "Please choose a valid input file first.")
            return

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

        # ====== PLUG YOUR BACKEND HERE ======
        # For now we run a simple parser as a placeholder:
        parsed = self.simple_convert(text)
        payload = {
            "sourceFile": str(self.current_input_path.name),
            "convertedAt": datetime.utcnow().isoformat() + "Z",
            "converter": "Quiz Converter (Desktop)",
            "summary": parsed.get("summary", ""),
            "data": parsed.get("data")
        }
        # ====================================

        try:
            out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
            self.show_toast(f"Converted and saved: {out_path.name}", 1800)
        except Exception as ex:
            QMessageBox.critical(self, "Save error", f"Failed to save output file:\n{ex}")

    def simple_convert(self, text: str):
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


# ------------ entrypoint ------------
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Quiz Converter")
    # ensure consistent QSS behavior on Windows
    app.setStyle("Fusion")

    window = QuizConverterMain()
    window.show()

    # center the window
    screen = app.primaryScreen().availableGeometry()
    geom = window.geometry()
    window.move((screen.width() - geom.width()) // 2, (screen.height() - geom.height()) // 2)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()