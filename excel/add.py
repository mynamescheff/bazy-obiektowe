#create a script to add random case number to all excel files in current folder, where the script is located. It should add "case number" into cells G1, then add random 10 alphanumeric characters to the cell.
import os
import random
import string
import openpyxl
from openpyxl import Workbook, load_workbook

# Function to generate a random alphanumeric string of length 10
def generate_random_string(length=10):
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(length))

# Get the current directory
current_directory = os.getcwd()

# Loop through all files in the current directory
for filename in os.listdir(current_directory):
    # Check if the file is an Excel file
    if filename.endswith('.xlsx') or filename.endswith('.xlsm'):
        # Load the workbook and select the active worksheet
        workbook = load_workbook(filename)
        sheet = workbook.active
        
        # Add "case number" to cell G1
        sheet['G1'] = 'case number'
        
        # Generate a random alphanumeric string and add it to cell G2
        random_string = generate_random_string()
        sheet['G2'] = random_string
        
        # Save the workbook with the changes
        workbook.save(filename)
        print(f"Updated {filename} with case number: {random_string}")
    else:
        print(f"Skipped {filename}, not an Excel file.")
# This script will add a "case number" header in cell G1 and a random alphanumeric string in cell G2 for each Excel file in the current directory.
# Make sure to have the openpyxl library installed. You can install it using pip:
