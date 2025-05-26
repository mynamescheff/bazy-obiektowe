import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook
import sqlite3

# This script converts specified text file into a .db file format.
# the text file looks as follows:
#--- Case List Created on 2025-05-26 17:44:39 ---

#--- Updated on 2025-05-26 17:44:39 ---
# XLxFodWStBq9vqp [Ella_Allen_University of Opole.xlsx] (2025-05-26 17:44:39)
#....
# --- Updated on 2025-05-26 17:44:56 ---
# so the script will read the text file, extract the relevant data, and save it to a SQLite database.


def convert_text_to_db():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    # Ask user to select a directory
    print("Please select a text file to convert to a database.")
    root.update()  # Update the root window to ensure it is ready for user interaction
    # Open a file dialog to select a text file

    # Ask user to select a text file
    file_path = filedialog.askopenfilename(title="Select a text file", filetypes=[("Text files", "*.txt")])
    # Check if a file was selected  
    if not file_path:
        print("No file selected.")
        return
    # Check if the selected file exists
    if not os.path.isfile(file_path):
        print(f"The file {file_path} does not exist.")
        return
    # Read the text file
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    # Extract relevant data from the text file
    data = []
    for line in lines:
        #start with the specific identifier to find relevant lines
        if line.startswith('case number: '):
            parts = line.split(' ')
            if len(parts) >= 3:
                filename = parts[1].strip('[]')
                timestamp = parts[2].strip('()')
                data.append({'filename': filename, 'timestamp': timestamp})
    if data:
        df = pd.DataFrame(data)
        # Create a SQLite database connection
        db_file_path = os.path.splitext(file_path)[0] + '.db'
        conn = sqlite3.connect(db_file_path)
        # Convert DataFrame to SQL table
        try:
            df.to_sql('data', conn, if_exists='replace', index=False)
        except Exception as e:
            print(f"Error converting DataFrame to SQL: {e}")
            conn.close()
            return
        # Close the database connection
        conn.close()
        print(f"Converted {file_path} to {db_file_path} successfully.")

if __name__ == "__main__":
    convert_text_to_db()