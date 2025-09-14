import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import customtkinter as ctk
import os
import time

class App(ctk.CTk):
    """
    A desktop application for a Quiz Converter with a modern UI.
    """

    def __init__(self):
        """Initializes the main application window and widgets with a new theme."""
        super().__init__()

        # Set appearance mode and default color theme
        ctk.set_appearance_mode("light")  # Can be "System", "Dark", or "Light"
        ctk.set_default_color_theme("blue")  # Can be "blue", "green", "dark-blue"

        # Window configuration
        self.title("Quiz Format Converter")
        self.geometry("600x800")

        # UI elements and styling
        self.selected_file_path = None
        self.converted_content = None
        self.title_font = ctk.CTkFont(family="Inter", size=24, weight="bold")
        self.subtitle_font = ctk.CTkFont(family="Inter", size=14)
        self.label_font = ctk.CTkFont(family="Inter", size=12)
        self.button_font = ctk.CTkFont(family="Inter", size=12, weight="bold")

        self.create_widgets()

    def create_widgets(self):
        """Creates the main UI components with improved styling."""
        
        # Header section
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", pady=(20, 10), padx=20)
        
        title_label = ctk.CTkLabel(header_frame, text="Quiz Format Converter", font=self.title_font, text_color="#1e293b")
        title_label.pack()
        
        subtitle_label = ctk.CTkLabel(header_frame, text="Convert quiz files from one format to another offline.", font=self.subtitle_font, text_color="#64748b")
        subtitle_label.pack(pady=(5, 0))

        # Main conversion section frame with a subtle shadow-like effect
        main_frame = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=10)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # File Upload Section
        upload_frame = ctk.CTkFrame(main_frame, fg_color="#f1f5f9", corner_radius=10)
        upload_frame.pack(fill="x", pady=(20, 10), padx=20)
        
        upload_label = ctk.CTkLabel(upload_frame, text="Upload a quiz file to get started.", font=self.label_font, text_color="#475569")
        upload_label.pack(pady=(10, 5))
        
        self.upload_btn = ctk.CTkButton(upload_frame, text="Browse Files", command=self.upload_file, font=self.button_font, fg_color="#dbeafe", text_color="#1d4ed8", hover_color="#bfdbfe")
        self.upload_btn.pack(pady=(0, 10), padx=40, fill="x")
        
        self.file_label = ctk.CTkLabel(upload_frame, text="No file selected", font=ctk.CTkFont(family="Inter", size=10), text_color="#94a3b8")
        self.file_label.pack(pady=(0, 10))
        
        # Convert Button
        self.convert_btn = ctk.CTkButton(main_frame, text="Convert File", command=self.convert_file, font=self.button_font, fg_color="#3b82f6", text_color="white", hover_color="#2563eb")
        self.convert_btn.pack(fill="x", pady=10, padx=20)
        self.convert_btn.configure(state="disabled", fg_color="#94a3b8")
        
        self.convert_status_label = ctk.CTkLabel(main_frame, text="", font=self.label_font, text_color="#3b82f6")
        self.convert_status_label.pack(pady=(5, 0))

        # Output/Preview section
        tk_label_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        tk_label_frame.pack(fill="x", pady=(20, 5), padx=20)

        tk_label = ctk.CTkLabel(tk_label_frame, text="Converted Output:", font=ctk.CTkFont(family="Inter", size=16, weight="bold"), text_color="#1f2937")
        tk_label.pack(anchor="w")
        
        self.output_text = ctk.CTkTextbox(main_frame, wrap="word", font=self.label_font, corner_radius=10, bg_color="#f1f5f9", fg_color="#f1f5f9", text_color="#1e293b", padx=10, pady=10)
        self.output_text.pack(fill="both", expand=True, padx=20)
        self.output_text.configure(state="disabled")
        
        # Save Button
        self.save_btn = ctk.CTkButton(main_frame, text="Save Converted File", command=self.save_file, font=self.button_font, fg_color="#4f46e5", text_color="white", hover_color="#4338ca")
        self.save_btn.pack(pady=20, padx=20, fill="x")
        self.save_btn.configure(state="disabled", fg_color="#94a3b8")
        
    def upload_file(self):
        """Opens a file dialog for the user to select a file."""
        file_path = filedialog.askopenfilename(
            title="Select a Quiz File",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
        )
        if file_path:
            self.selected_file_path = file_path
            self.file_label.configure(text=f"File: {os.path.basename(file_path)}")
            self.convert_btn.configure(state="normal", fg_color="#3b82f6")
            self.convert_status_label.configure(text="")
            self.output_text.configure(state="normal")
            self.output_text.delete("0.0", "end")
            self.output_text.configure(state="disabled")
            self.save_btn.configure(state="disabled", fg_color="#94a3b8")
            
    def convert_file(self):
        """Handles the file conversion logic and updates the UI."""
        if not self.selected_file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        self.convert_status_label.configure(text="Converting...")
        self.update()
        
        # Disable buttons during conversion
        self.convert_btn.configure(state="disabled", fg_color="#94a3b8")
        self.save_btn.configure(state="disabled", fg_color="#94a3b8")

        try:
            # --- Your specific Python conversion logic here ---
            with open(self.selected_file_path, 'r') as f:
                file_content = f.read()
            
            # This is a placeholder for your actual conversion code.
            self.converted_content = f"--- Converted Quiz Data ---\n\n{file_content}"
            
            time.sleep(1) # Simulate conversion time
            
            self.output_text.configure(state="normal")
            self.output_text.delete("0.0", "end")
            self.output_text.insert("end", self.converted_content)
            self.output_text.configure(state="disabled")
            
            self.convert_status_label.configure(text="Conversion complete!")
            self.save_btn.configure(state="normal", fg_color="#4f46e5")
            
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {e}")
            self.convert_status_label.configure(text="Conversion failed.", text_color="red")
        
        self.convert_btn.configure(state="normal", fg_color="#3b82f6")
            
    def save_file(self):
        """Opens a file dialog for the user to save the converted file."""
        if not self.converted_content:
            messagebox.showerror("Error", "No converted content to save.")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="Save Converted File As",
            defaultextension=".txt",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(self.converted_content)
                messagebox.showinfo("Success", "File saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
