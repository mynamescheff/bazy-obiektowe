import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook
import sqlite3

# This script converts specified Excel file into a .db file format.

def convert_excel_to_db():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    # Ask user to select a directory
    print("Please select an Excel file to convert to a database.")
    root.update()  # Update the root window to ensure it is ready for user interaction
    # Open a file dialog to select an Excel file

    # Ask user to select an Excel file
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not file_path:
        print("No file selected.")
        return

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Create a SQLite database connection
    db_file_path = os.path.splitext(file_path)[0] + '.db'
    conn = sqlite3.connect(db_file_path)
    
    # Convert DataFrame to SQL table
    df.to_sql('data', conn, if_exists='replace', index=False)

    # Close the database connection
    conn.close()

    print(f"Converted {file_path} to {db_file_path} successfully.")

if __name__ == "__main__":
    convert_excel_to_db()