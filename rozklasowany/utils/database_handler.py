import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class DatabaseHandler:

    def __init__(self, status_var=None):
        self.status_var = status_var

    def add_to_database(self):
        """
        Ask user to pick a .txt file, convert it to .db,
        then ask for an Excel file and convert it to .db.
        """
        # --- Process .txt file ---
        txt_file_path = filedialog.askopenfilename(
            title="Select a .txt file to convert to database",
            filetypes=[("Text files", "*.txt")]
        )
        if not txt_file_path:
            self._set_status("TXT file selection canceled.")
            messagebox.showinfo("Operation Canceled", "No .txt file selected. Skipping .txt conversion.")
        else:
            try:
                db_path_txt = self._convert_text_to_db(txt_file_path)
                messagebox.showinfo(
                    "Success",
                    f"Converted {os.path.basename(txt_file_path)} → {os.path.basename(db_path_txt)}"
                )
                self._set_status(f"TXT database saved as: {os.path.basename(db_path_txt)}")
            except Exception as e:
                messagebox.showerror("Error converting TXT", str(e))
                self._set_status(f"Error converting TXT: {e}")
            
        # --- Process Excel file ---
        excel_file_path = filedialog.askopenfilename(
            title="Select an Excel file to convert to database",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not excel_file_path:
            self._set_status("Excel file selection canceled.")
            messagebox.showinfo("Operation Canceled", "No Excel file selected. Skipping Excel conversion.")
        else:
            try:
                db_path_excel = self._convert_excel_to_db(excel_file_path)
                messagebox.showinfo(
                    "Success",
                    f"Converted {os.path.basename(excel_file_path)} → {os.path.basename(db_path_excel)}"
                )
                self._set_status(f"Excel database saved as: {os.path.basename(db_path_excel)}")
            except Exception as e:
                messagebox.showerror("Error converting Excel", str(e))
                self._set_status(f"Error converting Excel: {e}")

        if not txt_file_path and not excel_file_path:
            self._set_status("No files selected for conversion.")


    def _convert_text_to_db(self, txt_path: str) -> str:
        """
        Read the text file, extract lines starting 'case number: ',
        build a DataFrame with columns 'filename' and 'timestamp', then
        save to .db in the same directory.
        """
        self._set_status(f"Reading text file {os.path.basename(txt_path)}…")
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        records = []
        for line in lines:
            if line.startswith('case number: '):
                parts = line.strip().split()
                # e.g. ['case', 'number:', 'MyFile.xlsx', '(2025-05-26', '17:44:39)']
                if len(parts) >= 4:
                    filename = parts[2].strip('[]')
                    # join the rest for timestamp and strip parens
                    timestamp = ' '.join(parts[3:]).strip('()')
                    records.append({'filename': filename,
                                     'timestamp': timestamp})

        if not records:
            raise ValueError("No 'case number:' lines found in text file.")

        df = pd.DataFrame(records)
        db_path = os.path.splitext(txt_path)[0] + '.db'
        conn = sqlite3.connect(db_path)
        df.to_sql('data', conn, if_exists='replace', index=False)
        conn.close()
        return db_path

    def _convert_excel_to_db(self, excel_path: str) -> str:
        """
        Read the first sheet of the selected Excel file into a DataFrame
        and save to .db in the same directory.
        """
        self._set_status(f"Reading Excel file {os.path.basename(excel_path)}…")
        df = pd.read_excel(excel_path)
        if df.empty:
            raise ValueError("Selected Excel file contains no data.")

        db_path = os.path.splitext(excel_path)[0] + '.db'
        conn = sqlite3.connect(db_path)
        df.to_sql('data', conn, if_exists='replace', index=False)
        conn.close()
        return db_path

    def _set_status(self, text: str):
        if self.status_var is not None:
            self.status_var.set(text)