#create a script that scrapes the case numbers values from the excel files in the specified directory and saves the values in an existing excel file with the filename "case_number_db.xlsx".
#take values from column G "Case Number" for the scraping and put them into a column in the "case_number_db.xlsx" file, with the first column being "case number".
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook

def scrape_case_numbers():
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

    # Prepare a DataFrame to collect case number data
    case_data = []

    for file in excel_files:
        file_path = os.path.join(directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Check if 'case number' column exists
        if 'Case Number' in df.columns:
            for index, row in df.iterrows():
                case_number = row['Case Number']
                case_data.append({'Case Number': case_number})

    # Create a DataFrame from collected data
    case_df = pd.DataFrame(case_data)

    # Load or create the "case_number_db.xlsx" file
    db_file_path = os.path.join(directory, "case_number_db.xlsx")
    
    if os.path.exists(db_file_path):
        # Read existing data
        try:
            existing_df = pd.read_excel(db_file_path, sheet_name='Case Numbers')
        except Exception:
            existing_df = pd.DataFrame()
        # Concatenate old and new data
        combined_df = pd.concat([existing_df, case_df], ignore_index=True)
    else:
        combined_df = case_df

    # Save the combined DataFrame to the Excel file
    with pd.ExcelWriter(db_file_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='Case Numbers', index=False)

    print(f"Scraped case numbers and saved to {db_file_path} successfully.")

if __name__ == "__main__":
    scrape_case_numbers()