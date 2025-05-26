import os
import sqlite3
import pandas as pd
import tkinter as tk # Required for filedialog and messagebox if not already imported by main
from tkinter import filedialog, messagebox

# --- Constants for the verification function ---
# User should ensure these files exist and are populated correctly.
# Table name 'data' is assumed in both.
# combined.db expected columns: 'university', 'bank account' (and others)
# bank_acc_db.db expected columns: 'university', 'bank account'
COMBINED_DB_PATH_FOR_VERIFICATION = r".\\rozklasowany\\excelki\\cases\\combined\\combined.db"
BANK_ACC_DB_PATH_FOR_VERIFICATION = r".\\rozklasowany\\excelki\\bank_acc_db.db"

class DatabaseHandler:

    def __init__(self, status_var=None):
        self.status_var = status_var

    def _set_status(self, text: str):
        if self.status_var is not None:
            self.status_var.set(text)
        print(f"Status: {text}") # Also print to console for non-GUI contexts

    def add_to_database(self):
        """
        Ask user to pick a .txt file, convert it to .db,
        then ask for an Excel file and convert it to .db.
        This method uses filedialogs and is intended to be called from the GUI.
        """
        txt_file_path_selected = None 
        excel_file_path_selected = None

        # --- Process .txt file ---
        txt_file_path = filedialog.askopenfilename(
            title="Select a .txt file to convert to database",
            filetypes=[("Text files", "*.txt")]
        )
        if not txt_file_path:
            self._set_status("TXT file selection canceled.")
            messagebox.showinfo("Operation Canceled", "No .txt file selected. Skipping .txt conversion.")
        else:
            txt_file_path_selected = txt_file_path
            try:
                db_path_txt = self._convert_text_to_db(txt_file_path)
                messagebox.showinfo(
                    "Success",
                    f"Converted {os.path.basename(txt_file_path)} -> {os.path.basename(db_path_txt)}"
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
            excel_file_path_selected = excel_file_path
            try:
                db_path_excel = self._convert_excel_to_db(excel_file_path)
                messagebox.showinfo(
                    "Success",
                    f"Converted {os.path.basename(excel_file_path)} -> {os.path.basename(db_path_excel)}"
                )
                self._set_status(f"Excel database saved as: {os.path.basename(db_path_excel)}")
            except Exception as e:
                messagebox.showerror("Error converting Excel", str(e))
                self._set_status(f"Error converting Excel: {e}")

        if not txt_file_path_selected and not excel_file_path_selected:
            self._set_status("No files selected for conversion.")
        elif txt_file_path_selected or excel_file_path_selected:
             self._set_status("File to DB conversion process complete. Check messages for details.")


    def _convert_text_to_db(self, txt_path: str) -> str:
        """
        Read the text file, extract lines starting with 'case number: ',
        pull out the case_number (first token) and filename (second token),
        build a DataFrame with columns 'case_number' and 'filename', then
        save to .db in the same directory. Table name is 'data'.
        """
        self._set_status(f"Reading text file {os.path.basename(txt_path)}...")
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        records = []
        for line in lines:
            if line.lower().startswith('case number: '):
                # strip off the leading "case number:" and split
                content = line.strip()[len('case number: '):].strip()
                parts = content.split()
                # need at least two parts: case_number and filename
                if len(parts) >= 2:
                    case_number = parts[0]
                    filename = parts[1].strip('[]')
                    records.append({
                        'case_number': case_number,
                        'filename': filename
                    })

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
        and save to .db in the same directory. Table name is 'data'.
        """
        self._set_status(f"Reading Excel file {os.path.basename(excel_path)}...")
        df = pd.read_excel(excel_path) 
        if df.empty:
            raise ValueError("Selected Excel file contains no data.")

        db_path = os.path.splitext(excel_path)[0] + '.db'
        conn = sqlite3.connect(db_path)
        df.to_sql('data', conn, if_exists='replace', index=False)
        conn.close()
        return db_path

    def verify_bank_accounts_in_combined_db(self):
        """
        Verifies if 'bank account' values from 'combined.db' (table 'data')
        exist in 'bank_acc_db.db' (table 'data'), considering the 'university'.
        Paths to these databases are hardcoded constants at the top of this file.
        Uses self._set_status and messagebox for feedback.
        """
        self._set_status("Starting bank account verification...")
        
        required_cols_combined = ['university', 'bank account'] # Note the space
        required_cols_bank_acc = ['university', 'bank account'] # Note the space

        try:
            # --- Read bank_acc_db.db ---
            if not os.path.exists(BANK_ACC_DB_PATH_FOR_VERIFICATION):
                msg = f"Error: Database '{BANK_ACC_DB_PATH_FOR_VERIFICATION}' not found at expected path."
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                return

            conn_bank_acc = sqlite3.connect(BANK_ACC_DB_PATH_FOR_VERIFICATION)
            try:
                cursor_bank_acc = conn_bank_acc.cursor()
                cursor_bank_acc.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='data';")
                if not cursor_bank_acc.fetchone():
                    raise sqlite3.OperationalError(f"Table 'data' not found in {BANK_ACC_DB_PATH_FOR_VERIFICATION}")
                df_bank_acc = pd.read_sql_query("SELECT * FROM data", conn_bank_acc)
            except Exception as e: 
                msg = f"Error reading '{BANK_ACC_DB_PATH_FOR_VERIFICATION}': {e}"
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                conn_bank_acc.close()
                return
            finally:
                conn_bank_acc.close()

            if not all(col in df_bank_acc.columns for col in required_cols_bank_acc):
                msg = (f"Missing required columns in '{BANK_ACC_DB_PATH_FOR_VERIFICATION}'. "
                       f"Need: {', '.join(required_cols_bank_acc)}. Found: {', '.join(df_bank_acc.columns)}")
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                return
            
            bank_acc_set = set(zip(df_bank_acc['university'].astype(str), 
                                   df_bank_acc['bank account'].astype(str)))
            self._set_status(f"Loaded {len(bank_acc_set)} unique university/bank account pairs from {os.path.basename(BANK_ACC_DB_PATH_FOR_VERIFICATION)}.")

            # --- Read combined.db ---
            if not os.path.exists(COMBINED_DB_PATH_FOR_VERIFICATION):
                msg = f"Error: Database '{COMBINED_DB_PATH_FOR_VERIFICATION}' not found at expected path."
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                return

            conn_combined = sqlite3.connect(COMBINED_DB_PATH_FOR_VERIFICATION)
            try:
                cursor_combined = conn_combined.cursor()
                cursor_combined.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='data';")
                if not cursor_combined.fetchone():
                    raise sqlite3.OperationalError(f"Table 'data' not found in {COMBINED_DB_PATH_FOR_VERIFICATION}")
                df_combined = pd.read_sql_query("SELECT * FROM data", conn_combined)
            except Exception as e:
                msg = f"Error reading '{COMBINED_DB_PATH_FOR_VERIFICATION}': {e}"
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                conn_combined.close()
                return
            finally:
                conn_combined.close()
            
            if not all(col in df_combined.columns for col in required_cols_combined):
                msg = (f"Missing required columns in '{COMBINED_DB_PATH_FOR_VERIFICATION}'. "
                       f"Need: {', '.join(required_cols_combined)}. Found: {', '.join(df_combined.columns)}")
                self._set_status(msg)
                messagebox.showerror("Verification Error", msg)
                return

            not_found_accounts_info = []
            if df_combined.empty:
                self._set_status(f"{COMBINED_DB_PATH_FOR_VERIFICATION} is empty. No accounts to verify.")
                messagebox.showinfo("Verification Info", f"The database '{os.path.basename(COMBINED_DB_PATH_FOR_VERIFICATION)}' is empty.")
                return

            for index, row in df_combined.iterrows():
                uni = str(row['university'])
                acc = str(row['bank account']) # Note the space
                if (uni, acc) not in bank_acc_set:
                    record_info = f"Uni: {uni}, Acc: {acc}"
                    if 'filename' in df_combined.columns: record_info += f", File: {row.get('filename', 'N/A')}"
                    if 'name' in df_combined.columns and 'surname' in df_combined.columns:
                         record_info += f", Name: {row.get('name', 'N/A')} {row.get('surname', '')}"
                    not_found_accounts_info.append(record_info)

            if not_found_accounts_info:
                result_message = (f"Found {len(not_found_accounts_info)} bank account(s) from "
                                  f"'{os.path.basename(COMBINED_DB_PATH_FOR_VERIFICATION)}' that are NOT in "
                                  f"'{os.path.basename(BANK_ACC_DB_PATH_FOR_VERIFICATION)}':\n\n" +
                                  "\n".join(not_found_accounts_info[:10])) 
                if len(not_found_accounts_info) > 10:
                    result_message += f"\n\n...and {len(not_found_accounts_info) - 10} more. Check console for full list."
                
                self._set_status(f"Verification complete. {len(not_found_accounts_info)} accounts not found in {os.path.basename(BANK_ACC_DB_PATH_FOR_VERIFICATION)}.")
                messagebox.showwarning("Verification Result", result_message)
                print("\nFull list of accounts from combined.db not found in bank_acc_db.db:")
                for acc_info_item in not_found_accounts_info:
                    print(f"- {acc_info_item}")
            else:
                msg = (f"All {len(df_combined)} bank account(s) from '{os.path.basename(COMBINED_DB_PATH_FOR_VERIFICATION)}' "
                       f"were found in '{os.path.basename(BANK_ACC_DB_PATH_FOR_VERIFICATION)}'.")
                self._set_status(msg)
                messagebox.showinfo("Verification Result", msg)

        except Exception as e:
            error_msg = f"An unexpected error occurred during verification: {e}"
            self._set_status(error_msg)
            messagebox.showerror("Verification Error", error_msg)
            print(f"ERROR: {error_msg}")