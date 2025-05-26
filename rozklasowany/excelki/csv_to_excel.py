#ask user to select a csv file and convert it to an as many excel files as there are rows in the csv file, except the first row since it is the header
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os


def csv_to_excel():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Ask user to select a CSV file
    csv_file_path = filedialog.askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV files", "*.csv")]
    )

    if not csv_file_path:
        print("No file selected.")
        return

    # Read the CSV file
    df = pd.read_csv(csv_file_path)

    # Check if the DataFrame is empty or has only the header
    if df.empty or df.shape[0] <= 1:
        print("The CSV file is empty or has no data rows.")
        return

    # Create a directory for the Excel files
    output_dir = os.path.dirname(csv_file_path)
    
    for index, row in df.iterrows():
        if index == 0:  # Skip the header row
            continue
        
        # Create a new DataFrame for each row
        single_row_df = pd.DataFrame([row])
        
        # Generate a unique filename
        filename = f"{os.path.splitext(os.path.basename(csv_file_path))[0]}_row_{index}.xlsx"
        output_path = os.path.join(output_dir, filename)
        
        # Save the DataFrame to an Excel file
        single_row_df.to_excel(output_path, index=False)
    
    print(f"Converted {len(df) - 1} rows to Excel files in {output_dir}.")


if __name__ == "__main__":
    csv_to_excel()
