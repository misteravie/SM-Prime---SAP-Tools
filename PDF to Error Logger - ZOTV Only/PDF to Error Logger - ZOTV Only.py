# -*- coding: utf-8 -*-
"""
Created on Mon Jul 15 13:53:19 2024

@author: sanch
"""



import re
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import logging

# Configure logging
logging.basicConfig(filename='transaction_log_processor.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def parse_log_file(file_path):
    """Parse the log file and extract unsuccessful transactions."""
    transactions_with_errors = []

    try:
        with fitz.open(file_path) as doc:
            content = ""
            for page in doc:
                content += page.get_text()

        transactions = content.split("Transaction Number:")
        for transaction in transactions:
            if not transaction.strip():
                continue

            lines = transaction.splitlines()
            transaction_number = lines[0].strip()
            
            # Check if transaction contains "Error in document:"
            if any("Error in document:" in line for line in lines):
                error_logs = []
                for line in lines[1:]:
                    if line.strip():
                        error_logs.append(line.strip())
                transactions_with_errors.append((transaction_number, error_logs))

    except Exception as e:
        logging.error(f"Failed to parse log file {file_path}: {e}")
        raise

    return transactions_with_errors

def write_to_xlsx(all_transactions, output_file):
    """Write the unsuccessful transactions to a single XLSX file."""
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Unsuccessful Transactions"

        # Write the header
        header = ['File Name', 'Transaction Number', 'Error Log 1', 'Error Log 2', 'Error Log 3']
        ws.append(header)

        for file_name, transaction_number, error_logs in all_transactions:
            # Ensure the row has enough columns (3 for error logs)
            row = [file_name, transaction_number] + error_logs[:3] + [''] * (3 - len(error_logs[:3]))
            ws.append(row)

        wb.save(output_file)
        logging.info(f"Successfully wrote data to {output_file}")
    except Exception as e:
        logging.error(f"Failed to write to XLSX file {output_file}: {e}")
        raise

def process_files(input_dir):
    """Process all PDF files in the input directory."""
    all_transactions = []

    try:
        for filename in os.listdir(input_dir):
            if filename.endswith('.pdf'):
                input_file_path = os.path.join(input_dir, filename)
                transactions_with_errors = parse_log_file(input_file_path)
                all_transactions.extend([(filename, trans_num, error_logs) for trans_num, error_logs in transactions_with_errors])
                logging.info(f"Processed {filename}")
    except Exception as e:
        logging.error(f"Failed to process files in directory {input_dir}: {e}")
        raise

    return all_transactions

def run_script(input_dir, output_file):
    """Run the main script."""
    try:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        all_transactions = process_files(input_dir)
        write_to_xlsx(all_transactions, output_file)
        messagebox.showinfo("Success", f"All transactions have been written to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        logging.error(f"An error occurred during script execution: {e}")

class Application(tk.Frame):
    """GUI Application for the transaction log processor."""
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.input_label = tk.Label(self, text="Input Directory:")
        self.input_label.grid(row=0, column=0, padx=10, pady=10)

        self.input_entry = tk.Entry(self, width=50)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10)

        self.input_button = tk.Button(self, text="Browse", command=self.browse_input)
        self.input_button.grid(row=0, column=2, padx=10, pady=10)

        self.output_label = tk.Label(self, text="Output File:")
        self.output_label.grid(row=1, column=0, padx=10, pady=10)

        self.output_entry = tk.Entry(self, width=50)
        self.output_entry.grid(row=1, column=1, padx=10, pady=10)

        self.output_button = tk.Button(self, text="Browse", command=self.browse_output)
        self.output_button.grid(row=1, column=2, padx=10, pady=10)

        self.run_button = tk.Button(self, text="Run", command=self.run)
        self.run_button.grid(row=2, column=1, padx=10, pady=10)

    def browse_input(self):
        input_dir = filedialog.askdirectory()
        self.input_entry.insert(0, input_dir)

    def browse_output(self):
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        self.output_entry.insert(0, output_file)

    def run(self):
        input_dir = self.input_entry.get()
        output_file = self.output_entry.get()
        run_script(input_dir, output_file)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Transaction Log Processor")
    app = Application(master=root)
    app.mainloop()
