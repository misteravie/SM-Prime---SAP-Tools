# -*- coding: utf-8 -*-
"""
Created on Fri Jun  7 13:26:44 2024

@author: sanch
"""

import re
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF

# Function to parse the log file and extract successful transactions
def parse_log_file(file_path):
    transactions_data = []

    # Open the PDF file
    with fitz.open(file_path) as doc:
        content = ""
        for page in doc:
            content += page.get_text()

    # Split content by transactions
    transactions = content.split("Transaction Number:")
    for transaction in transactions:
        if not transaction.strip():
            continue
        
        lines = transaction.splitlines()
        transaction_number = lines[0].strip()
        
        for line in lines[1:]:
            if line.startswith('S ') and "was posted" in line:
                success_log = line.strip()
                transactions_data.append((transaction_number, success_log))
                break
        else:
            transactions_data.append((transaction_number, ""))  # Append empty success log if none found

    return transactions_data

# Function to write the transactions to a single XLSX file
def write_to_xlsx(all_transactions, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    # Write the header
    ws.append(['File Name', 'Transaction Number', 'Success Log'])

    # Write the data
    for file_name, transaction_number, success_log in all_transactions:
        ws.append([file_name, transaction_number, success_log])

    wb.save(output_file)

# Function to process all PDF files in the input directory
def process_files(input_dir):
    all_transactions = []

    # Process each PDF file in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.pdf'):
            input_file_path = os.path.join(input_dir, filename)
            transactions_data = parse_log_file(input_file_path)
            # Add the filename to each transaction
            all_transactions.extend([(filename, trans_num, success_log) for trans_num, success_log in transactions_data])
            print(f"Processed {filename}")

    return all_transactions

# Function to run the script
def run_script(input_dir, output_file):
    try:
        # Ensure the output directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        all_transactions = process_files(input_dir)
        write_to_xlsx(all_transactions, output_file)
        messagebox.showinfo("Success", f"All transactions have been written to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI Application
class Application(tk.Frame):
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

root = tk.Tk()
root.title("Transaction Log Processor")
app = Application(master=root)
app.mainloop()
