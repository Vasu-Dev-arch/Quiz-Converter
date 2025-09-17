import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class QuizFormatterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("QuizFormatter")
        self.selected_theme = tk.StringVar(value="Light")
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.set_theme("Light")  # default theme

        # ---- Top bar: Title and Settings icon ----
        top_bar = ttk.Frame(root, style="Toolbar.TFrame")
        top_bar.pack(fill="x", padx=0, pady=0)
        title_label = ttk.Label(top_bar, text="QuizFormatter", style="Toolbar.TLabel", font=("Segoe UI", 16, "bold"))
        title_label.pack(side="left", padx=(18, 6), pady=6)
        
        # Settings icon using built-in bitmap (no Pillow needed)
        settings_icon = tk.Label(top_bar, bitmap="questhead", cursor="hand2", background="#ECEFF4")
        settings_icon.pack(side="right", padx=16, pady=2)
        settings_icon.bind("<Button-1>", self.show_settings_menu)
        
        # ---- Main Card Panel ----
        main_panel = ttk.Frame(root, style="Card.TFrame", padding=32)
        main_panel.place(relx=0.5, rely=0.5, anchor="center")

        # Input row
        ttk.Label(main_panel, text="Input File (.docx):", style="Card.TLabel", anchor="w").grid(row=0, column=0, sticky="w")
        self.input_entry = ttk.Entry(main_panel, textvariable=self.input_path, width=44, state="readonly", font=("Segoe UI", 10))
        self.input_entry.grid(row=1, column=0, sticky="ew", pady=(0,12))
        self.browse_btn = ttk.Button(main_panel, text="Browse", style="Accent.TButton", command=self.browse_input)
        self.browse_btn.grid(row=1, column=1, padx=(12,0))

        # Output row
        ttk.Label(main_panel, text="Output File (.docx):", style="Card.TLabel", anchor="w").grid(row=2, column=0, sticky="w", pady=(16,0))
        self.output_entry = ttk.Entry(main_panel, textvariable=self.output_path, width=44, state="readonly", font=("Segoe UI", 10))
        self.output_entry.grid(row=3, column=0, sticky="ew", pady=(0,12))
        self.saveas_btn = ttk.Button(main_panel, text="Save As", style="Accent.TButton", command=self.browse_output)
        self.saveas_btn.grid(row=3, column=1, padx=(12,0))

        # Convert button
        self.convert_btn = ttk.Button(main_panel, text="Convert", style="Accent.TButton", command=self.convert)
        self.convert_btn.grid(row=4, column=0, columnspan=2, pady=(30,0), ipadx=18, ipady=6)

        main_panel.columnconfigure(0, weight=1)

    def show_settings_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        theme_menu = tk.Menu(menu, tearoff=0)
        theme_menu.add_radiobutton(label="Light", variable=self.selected_theme, command=lambda:self.set_theme("Light"))
        theme_menu.add_radiobutton(label="Dark", variable=self.selected_theme, command=lambda:self.set_theme("Dark"))
        menu.add_cascade(label="Theme", menu=theme_menu)
        menu.tk_popup(event.x_root, event.y_root)
        
    def set_theme(self, theme):
        s = ttk.Style()
        if theme == "Light":
            self.root.configure(bg="#EFF2F7")
            s.configure('Card.TFrame', background="#fff", borderwidth=1, relief="ridge")
            s.configure('Accent.TButton', background="#1976D2", foreground="white", font=("Segoe UI", 10, "bold"))
            s.map('Accent.TButton', background=[('active', '#1565c0')])
            s.configure('Card.TLabel', background="#fff", font=("Segoe UI", 12))
            s.configure('Toolbar.TFrame', background="#ECEFF4")
            s.configure('Toolbar.TLabel', background="#ECEFF4", foreground="#222")
        else:
            # Dark mode colors
            self.root.configure(bg="#262B38")
            s.configure('Card.TFrame', background="#232634", borderwidth=0)
            s.configure('Accent.TButton', background="#5EA6F7", foreground="white", font=("Segoe UI", 10, "bold"))
            s.map('Accent.TButton', background=[('active', '#0066cc')])
            s.configure('Card.TLabel', background="#232634", foreground="#eee", font=("Segoe UI", 12))
            s.configure('Toolbar.TFrame', background="#232634")
            s.configure('Toolbar.TLabel', background="#232634", foreground="#eaeaea")
        
    def browse_input(self):
        fn = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if fn:
            self.input_path.set(fn)

    def browse_output(self):
        fn = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if fn:
            self.output_path.set(fn)
    
    def convert(self):
        # Stub -- add your conversion logic here
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showwarning("Missing files", "Please select both input and output files.")
            return
        messagebox.showinfo("Ready!", f"Would convert:\nInput: {self.input_path.get()}\nOutput: {self.output_path.get()}")

if __name__ == '__main__':
    root = tk.Tk()
    root.geometry('450x430')
    app = QuizFormatterGUI(root)
    root.mainloop()
