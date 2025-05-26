#append column with case number to excel files in the directory with the filename starting with "ludzie", include the case number in the cell row below the header, and save the file with the values from column A + B + C.
#make the case number be a random string of letters and numbers, 15 characters long
import tkinter as tk
from tkinter import filedialog
import os
import re
import pandas as pd
import random
import string

def generate_case_number(length=15):
    """Generate a random case number consisting of letters and digits."""
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(length))

def append_case_number_to_excel_files():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Ask user to select a directory
    directory = filedialog.askdirectory(title="Select a directory containing Excel files")
    
    if not directory:
        print("No directory selected.")
        return

    # List all Excel files in the directory
    excel_files = [f for f in os.listdir(directory) if f.startswith("ludzie") and f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("No Excel files starting with 'ludzie' found in the selected directory.")
        return

    for file in excel_files:
        file_path = os.path.join(directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Generate a case number
        case_number = generate_case_number()

        # Append the case number to the DataFrame
        df.loc[0, 'Case Number'] = case_number  # Add case number below the header

        # Create a new filename based on columns A, B, C values
        new_filename = f"{df.iloc[0, 0]}_{df.iloc[0, 1]}_{df.iloc[0, 2]}.xlsx"
        new_file_path = os.path.join(directory, new_filename)

        # Save the modified DataFrame to a new Excel file
        df.to_excel(new_file_path, index=False)

    print(f"Case numbers appended and files saved in {directory}.")

if __name__ == "__main__":
    append_case_number_to_excel_files()
# This script appends a random case number to Excel files in a specified directory,
