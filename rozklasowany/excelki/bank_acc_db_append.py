#create a script that scrapes the bank account values from the excel files in the directory starting with "ludzie", and saves the values in an existing excel file with the filename "bank_acc_db.xlsx".
#take values from column C "university"	and from column D "bank account" for the scraping and put them into two columns in the "bank_acc_db.xlsx" file, with the first column being "university" and the second column being "bank account".

import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook

def scrape_bank_accounts():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Ask user to select a directory
    directory = filedialog.askdirectory(title="Select a directory containing Excel files")
    
    if not directory:
        print("No directory selected.")
        return

    # List all Excel files in the directory starting with "ludzie"
    excel_files = [f for f in os.listdir(directory) if f.startswith("ludzie") and f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("No Excel files starting with 'ludzie' found in the selected directory.")
        return

    # Prepare a DataFrame to collect bank account data
    bank_data = []

    for file in excel_files:
        file_path = os.path.join(directory, file)
        
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Check if required columns exist
        if 'university' in df.columns and 'bank account' in df.columns:
            for index, row in df.iterrows():
                university = row['university']
                bank_account = row['bank account']
                bank_data.append({'university': university, 'bank account': bank_account})

    # Create a DataFrame from collected data
    bank_df = pd.DataFrame(bank_data)

    # Load or create the "bank_acc_db.xlsx" file
    db_file_path = os.path.join(directory, "bank_acc_db.xlsx")
    
    if os.path.exists(db_file_path):
        # Read existing data
        try:
            existing_df = pd.read_excel(db_file_path, sheet_name='Bank Accounts')
        except Exception:
            existing_df = pd.DataFrame()
        # Concatenate old and new data
        combined_df = pd.concat([existing_df, bank_df], ignore_index=True)
    else:
        combined_df = bank_df

    # Write the combined DataFrame to the Excel file
    with pd.ExcelWriter(db_file_path, engine='openpyxl', mode='w') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Bank Accounts')

    print(f"Bank account data scraped and saved to {db_file_path}.")

if __name__ == "__main__":
    scrape_bank_accounts()