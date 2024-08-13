import tkinter as tk
from tkinter import filedialog, messagebox
from fpdf import FPDF
import os


#This is a test

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text to PDF Converter")
        
        self.label = tk.Label(root, text="Select Text Files to Convert to PDF")
        self.label.pack(pady=10)
        
        self.convert_button = tk.Button(root, text="Select Files", command=self.select_files)
        self.convert_button.pack(pady=10)
        
        self.quit_button = tk.Button(root, text="Quit", command=root.quit)
        self.quit_button.pack(pady=10)
        
        self.files = []
        
    def select_files(self):
        self.files = filedialog.askopenfilenames(title="Select Text Files", filetypes=(("Text Files", "*.txt"),))
        if self.files:
            self.convert_to_pdf()
    
    def convert_to_pdf(self):
        for file in self.files:
            try:
                self.create_pdf(file)
                messagebox.showinfo("Success", f"Successfully converted {os.path.basename(file)} to PDF")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert {os.path.basename(file)} to PDF\n{str(e)}")
                
    def create_pdf(self, txt_file):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        
        with open(txt_file, 'r', encoding='utf-8') as file:
            for line in file:
                pdf.cell(200, 10, txt=line, ln=True)
        
        output_file = os.path.splitext(txt_file)[0] + ".pdf"
        pdf.output(output_file)
        
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()