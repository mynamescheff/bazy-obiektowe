#change column G to "case number" in every excel file in the directory selected by the user and save the file

import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook

def change_case_number_column():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Ask user to select a directory
    directory = filedialog.askdirectory(title="Select a directory containing Excel files")
    
    if not directory:
        print("No directory selected.")
        return

    # List all Excel files in the directory
    excel_files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("No Excel files found in the selected directory.")
        return

    for file in excel_files:
        file_path = os.path.join(directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Check if 'Case Number' column exists
        if 'Case Number' in df.columns:
            # Rename the column to 'case number'
            df.rename(columns={'Case Number': 'case number'}, inplace=True)

            # Save the modified DataFrame back to the Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            print(f"Changed 'Case Number' to 'case number' in {file}")

        else:
            print(f"'Case Number' column not found in {file}, skipping.")

if __name__ == "__main__":
    change_case_number_column()