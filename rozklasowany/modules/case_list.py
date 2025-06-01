import os
import re
from datetime import datetime
from openpyxl import load_workbook

class CaseList:
    def __init__(self, excel_folder, list_folder):
        self.excel_folder = excel_folder
        self.list_folder = list_folder

    def process_excel_files(self, text_widget_update):
        list_file_path = os.path.join(self.list_folder, "case_list.txt")
        existing_values = {}
        duplicate_counts = {}
        error_messages = []
        
        if os.path.isfile(list_file_path):
            existing_values, duplicate_counts = self.load_existing_list(list_file_path)
            
        all_entries = []
        for file_name in os.listdir(self.excel_folder):
            if file_name.endswith(".xlsx"):
                file_path = os.path.join(self.excel_folder, file_name)
                try:
                    wb = load_workbook(file_path)
                    sheet = wb.active
                    value = sheet["G2"].value
                    if value:
                        # Clean the value
                        value = self._clean_string(value)
                        
                    if value in existing_values:
                        duplicate_counts[value] = duplicate_counts.get(value, 0) + 1
                        entry = f"case number: {value} [{file_name} - DUPLICATE {duplicate_counts[value]}]"
                        all_entries.append(entry)
                        # Log the duplicate in the gui place in the app
                        text_widget_update(f"Duplicate case number found: {value} in file {file_name}. Count: {duplicate_counts[value]}")
                    elif value:
                        existing_values[value] = True
                        duplicate_counts[value] = 0
                        entry = f"case number: {value} [{file_name}]"
                        all_entries.append(entry)
                        # Log the new case number in the gui place in the app
                        text_widget_update(f"New case number added: {value} from file {file_name}.")

                    else:
                        existing_values[value] = False
                        duplicate_counts[value] = 0
                        entry = f"case number: {value} [{file_name}]"
                        all_entries.append(entry)
                except Exception as e:
                    error_messages.append(f"Error processing file '{file_name}': {str(e)}")
                    
        if all_entries:
            today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Ensure the directory exists
            os.makedirs(os.path.dirname(list_file_path), exist_ok=True)
            
            # Create the file if it doesn't exist
            if not os.path.exists(list_file_path):
                with open(list_file_path, "w", encoding="utf-8") as file:
                    file.write(f"--- Case List Created on {today} ---\n")
            
            # Append the new entries
            with open(list_file_path, "a", encoding="utf-8") as file:
                file.write(f"\n--- Updated on {today} ---\n")
                for entry in all_entries:
                    file.write(f"{entry} ({today})\n")

                text_widget_update("Cases processed successfully and saved to the case_list.txt file.")
                return duplicate_counts, error_messages

    def _clean_string(self, value):
        if not value:
            return ""
            
        value = str(value)
        value = re.sub(r'[\n\r\t\v\f\x85\u2028\u2029]+', ' ', value)
        
        # Remove extra spaces
        while "  " in value:
            value = value.replace("  ", " ")
            
        # Make sure string ends with alphanumeric character
        while value and not value[-1].isalnum():
            value = value[:-1]
            
        return value.strip()

    def load_existing_list(self, list_file_path):
        existing_values = {}
        duplicate_counts = {}
        with open(list_file_path, "r", encoding="utf-8") as file:
            for line in file:
                if line.strip() and not line.startswith("---"):
                    match = re.match(r"case number: (.+?) \[", line)
                    if match:
                        value = match.group(1)
                        existing_values[value] = False
                        duplicate_count = line.count("DUPLICATE")
                        duplicate_counts[value] = duplicate_count
        return existing_values, duplicate_counts
