import os
import re
import time
import shutil
import datetime
from datetime import date, datetime
from pathlib import Path
from tkinter import Tk, Frame, Label, Button, Entry, Text, Toplevel, filedialog, messagebox, Scrollbar, RIGHT, Y, Checkbutton, BooleanVar, W, E
from tkinter import *
import tkinter as tk
import win32com.client
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import queue
import threading
import sys
import pythoncom
import csv
import openpyxl
from openpyxl.cell.cell import MergedCell

SHARED_MAILBOX_EMAIL = "" #insert email here
notes_window = None
instructions_window = None
notes_content = ""
notes_first_time = True

class CharacterTransformer:
    def __init__(self):
        self.character_mapping = {
            'á': 'a', 'à': 'a', 'â': 'a', 'ä': 'a', 'ã': 'a', 'å': 'a', 'æ': 'ae', 'ç': 'c', 'č': 'c', 'ć': 'c',
            'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e', 'ė': 'e', 'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i', 'ñ': 'n',
            'ó': 'o', 'ò': 'o', 'ô': 'o', 'ö': 'o', 'õ': 'o', 'ø': 'o', 'œ': 'oe', 'š': 's', 'ß': 'ss', 'ú': 'u',
            'ù': 'u', 'û': 'u', 'ü': 'u', 'ý': 'y', 'ÿ': 'y', 'ž': 'z', 'Á': 'A', 'À': 'A', 'Â': 'A', 'Ä': 'A',
            'Ã': 'A', 'Å': 'A', 'Æ': 'AE', 'Ç': 'C', 'Č': 'C', 'Ć': 'C', 'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
            'Ė': 'E', 'Í': 'I', 'Ì': 'I', 'Î': 'I', 'Ï': 'I', 'Ñ': 'N', 'Ó': 'O', 'Ò': 'O', 'Ô': 'O', 'Ö': 'O',
            'Õ': 'O', 'Ø': 'O', 'Œ': 'OE', 'Š': 'S', 'Ú': 'U', 'Ù': 'U', 'Û': 'U', 'Ü': 'U', 'Ý': 'Y', 'Ÿ': 'Y',
            'Ž': 'Z', 'Þ': 'th', 'þ': 'th', 'ð': 'dh', 'Đ': 'D', 'đ': 'd', 'ł': 'l', 'Ł': 'L', 'ū': 'u', 'Ū': 'U',
            'Ā': 'A', 'ā': 'a', 'Ē': 'E', 'ē': 'e', 'Ī': 'I', 'ī': 'i', 'Ō': 'O', 'ō': 'o', 'ă': 'a', 'Ă': 'A',
            'ș': 's', 'Ș': 'S', 'ț': 't', 'Ț': 'T', 'â': 'a', 'Â': 'A', 'î': 'i', 'Î': 'I', 'ş': 's', 'Ş': 'S',
            'ğ': 'g', 'Ğ': 'G', 'İ': 'I', 'ı': 'i', 'ö': 'o', 'Ö': 'O', 'ü': 'u', 'Ü': 'U', 'ț': 't', 'Ţ': 'T', 'ţ': 't',
            
            '_': ' ', '"': ' ', '  ': ' ', '\xa0': ' ', '\t': ' ', '\n': ' ', '\r': ' ', '\x0b': ' ', '\x0c': ' ',
            '\u200b': ' ', '\u200c': ' ', '\u200d': ' ', '\u200e': ' ', '\u200f': ' ', '\u202a': ' ', '\u202c': ' ',
            '\u202d': ' ', '\u202e': ' ', '\u202f': ' ', '\u205f': ' ', '\u3000': ' ', '\u2000': ' ', '\u2001': ' ',
            '\u2002': ' ', '\u2003': ' ', '\u2004': ' ', '\u2005': ' ', '\u2006': ' ', '\u2007': ' ', '\u2008': ' ',
            '\u2009': ' ', '\u200a': ' ', ' ': ' ', ' ': ' ', '&nbsp;':'', '\u0219' : 's', 'ș': 's', 'ă': 'a', 'ț': 't', 
            'ő': 'o', 'Ő': 'O', 'ű': 'u', 'Ű': 'U', '\n': ' ', '\r': ' ', '\t': ' ', '\xa0': ' ', '\u200b': ' ', '\u200c': ' ',
            'ę': 'e', 'Ę': 'E', 'ą': 'a', 'Ą': 'A', 'ś': 's', 'Ś': 'S', 'ł': 'l', 'Ł': 'L', 'ż': 'z', 'Ż': 'Z', 'ź': 'z', 'Ź': 'Z',
            ';': '', 'ń':'n',
        }

    def transform_to_swift_accepted_characters(self, input_list):
        transformed_list = []
        for input_string in input_list:
            transformed_string = re.sub(r'\b\w+\b', lambda m: ''.join(self.character_mapping.get(char, char) for char in m.group()), str(input_string))
            transformed_string = re.sub(r'[.,]', '', transformed_string)
            transformed_list.append(transformed_string)
        return transformed_list

class Wide:
    def __init__(self, file_path: str, sheet_name: str, directory: str):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.directory = directory
        self.workbook = load_workbook(file_path)
        self.sheet = self.workbook[sheet_name]

    def auto_adjust_column_width(self) -> None:
        for column in self.sheet.columns:
            max_length = 0
            column_cells = [cell for cell in column]
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            self.sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = adjusted_width
        self.workbook.save(self.file_path)
        print("Column widths adjusted and saved successfully.")

    def get_file_count(self) -> int:
        file_count = len([name for name in os.listdir(self.directory) if os.path.isfile(os.path.join(self.directory, name))])
        return file_count

    def create_table_with_headers(self, table_name: str) -> None:
        file_count = self.get_file_count()
        if file_count == 0:
            print("No files found in the directory.")
            return
        last_row = file_count + 1
        last_column = self.sheet.max_column
        last_column_letter = get_column_letter(last_column)
        new_sheet = self.workbook.create_sheet(title=f"{table_name}_Sheet")
        for row in self.sheet.iter_rows(max_row=last_row, max_col=last_column):
            new_sheet.append([cell.value for cell in row])
        self.auto_adjust_column_width(new_sheet)
        table_range = f"A1:{last_column_letter}{last_row}"
        table = Table(displayName=table_name, ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        table.tableStyleInfo = style
        new_sheet.add_table(table)
        self.workbook.save(self.file_path)
        print("Table with headers created in new sheet and saved successfully.")

class ExcelComparator:
    def __init__(self, combined_file, sprawdzacz_file, output_path):
        self.combined_file = combined_file
        self.sprawdzacz_file = sprawdzacz_file
        self.output_path = output_path

    def compare_and_append(self):
        try:
            combined_workbook = load_workbook(filename=self.combined_file)
            combined_sheet = combined_workbook['Transposed']
            sprawdzacz_workbook = load_workbook(filename=self.sprawdzacz_file)
            sprawdzacz_sheet = sprawdzacz_workbook['name_acc']
            non_matching_values = []
            filenames = []

            # Check and rename duplicates in column B
            b_values = {}
            for row in range(2, combined_sheet.max_row + 1):
                b_value = combined_sheet[f'B{row}'].value

                if b_value in b_values:
                    b_values[b_value].append(row)
                else:
                    b_values[b_value] = [row]
            
            for b_value, rows in b_values.items():
                if len(rows) > 1:
                    counter = 2  # Start numbering from 2
                    for row in rows:
                        a_value = combined_sheet[f'A{row}'].value.replace(" ", "")
                        if any(char.isdigit() for char in a_value):
                            combined_sheet[f'B{row}'].value = f"{b_value}{counter}"
                            counter += 1

            # Save the workbook after renaming
            combined_workbook.save(self.combined_file)

            # Compare values as before
            for row in range(2, combined_sheet.max_row + 1):
                case_value = combined_sheet[f'A{row}'].value
                f_value = str(combined_sheet[f'F{row}'].value).replace(" ", "").replace("-", "") if combined_sheet[f'F{row}'].value else ""
                h_value = str(combined_sheet[f'G{row}'].value).replace(" ", "").replace("-", "") if combined_sheet[f'G{row}'].value else ""
                found_match = any(f_value == str(row_val[2]) or h_value == str(row_val[2]) for row_val in sprawdzacz_sheet.iter_rows(values_only=True))
                if not found_match:
                    non_matching_values.append((case_value, f_value, h_value))
                    filenames.append((self.combined_file, self.sprawdzacz_file))

            file_path = os.path.join(self.output_path, 'utils\\mismatch_list.txt')
            with open(file_path, 'w', encoding="utf-8") as file:
                current_date = date.today()
                for value in non_matching_values:
                    modified_value = value[0].replace("\xa0", "").strip()
                    file.write(f"{modified_value}: {value[1]}, {value[2]} ({current_date})\n")

            return non_matching_values, f"Non-matching values appended to {file_path} successfully."
        except Exception as e:
            return [], str(e)


class CaseList:
    def __init__(self, excel_folder, list_folder):
        self.excel_folder = excel_folder
        self.list_folder = list_folder

    def process_excel_files(self):
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
                    value = sheet["B17"].value
                    if value in existing_values:
                        value.replace(" ", "").replace("\n", "").replace("\t", "").replace("\r", "").replace("  ", " ").replace(" ", "").replace("\xa0", "").replace(" ", "").lstrip().replace("\n", " ").replace("\t", " ").replace("\r", " ").replace("\x0b", " ").replace("\x0c", " ").replace("\t", " ").replace("\r", " ").replace("\x85", " ").replace("\r\n", " ").replace("\u2028"," ").replace("\u2029"," ")
                        value.replace(" ", "")
                        #make sure that the string of the value variable ends on a number or a letter
                        while not value[-1].isalnum():
                            value = value[:-1]
                        value = re.sub(r'[\n\r\t\v\f\x85\u2028\u2029]+', ' ', value)
                        #keep replacing double spaces from the string until there are only singular space overall
                        while "  " in value:
                            value = value.replace("  ", " ")
                        while "\n" in value:
                            value = value.replace("\n", "")
                        duplicate_counts[value] += 1
                        existing_values[value] = True
                        entry = f"{value} [{file_name} - DUPLICATE {duplicate_counts[value]}]"
                        all_entries.append(entry)
                        print(f"Duplicate found: {value} in file {file_name} (Duplicate count: {duplicate_counts[value]})")
                    else:
                        existing_values[value] = False
                        duplicate_counts[value] = 0
                        entry = f"{value} [{file_name}]"
                        all_entries.append(entry)
                except Exception as e:
                    error_messages.append(f"Error processing file '{file_name}': {str(e)}")
        if all_entries:
            today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(list_file_path, "a", encoding="utf-8") as file:
                file.write(f"\n--- Updated on {today} ---\n")
                for entry in all_entries:
                    file.write(f"{entry} ({today})\n")
        print("Cases processed successfully and saved to the case_list.txt file.")
        return duplicate_counts, error_messages

    def load_existing_list(self, list_file_path):
        existing_values = {}
        duplicate_counts = {}
        with open(list_file_path, "r", encoding="utf-8") as file:
            for line in file:
                if line.strip() and not line.startswith("---"):
                    parts = line.split(" [")
                    value = parts[0]
                    existing_values[value] = False
                    duplicate_count = parts[1].count("DUPLICATE")
                    duplicate_counts[value] = duplicate_count
        return existing_values, duplicate_counts

class OutlookProcessor:
    def __init__(self, category, target_senders, attachment_save_path, msg_save_path):
        self.category = category
        self.target_senders = target_senders
        self.attachment_save_path = attachment_save_path
        self.msg_save_path = msg_save_path
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        os.makedirs(self.attachment_save_path, exist_ok=True)
        os.makedirs(self.msg_save_path, exist_ok=True)

    def list_unread_emails(self):
        recipient = self.namespace.CreateRecipient(SHARED_MAILBOX_EMAIL)
        recipient.Resolve()
        if recipient.Resolved:
            shared_mailbox = self.namespace.GetSharedDefaultFolder(recipient, 6)
            unread_emails = shared_mailbox.Items.Restrict(f"[Categories] = '{self.category}' AND [UnRead] = True")
            return len([email for email in unread_emails])
        return 0

    def get_unique_filename(self, base_path, original_filename, extension):
        counter = 2
        new_filename = original_filename
        while os.path.exists(os.path.join(base_path, f"{new_filename}{extension}")):
            new_filename = f"{original_filename} {counter}"
            counter += 1
        return new_filename

    def mark_email_as_read(self, item, mark_as_read):
        item.UnRead = not mark_as_read
        item.Save()
        if mark_as_read:
            print("Marked email as read.")
        else:
            print("Email left as unread.")

    def download_attachments_and_save_as_msg(self, save_emails, mark_as_read):
        recipient = self.namespace.CreateRecipient(SHARED_MAILBOX_EMAIL)
        recipient.Resolve()
        self.processed_emails = {}  # Clear previous tracking
        self.emails_with_pdf = []  # Clear previous tracking
        self.emails_with_nvf_new_vendor = []  # Clear previous tracking
        emails_with_no_attachments = []  # Track emails with no attachments
        if recipient.Resolved:
            shared_mailbox = self.namespace.GetSharedDefaultFolder(recipient, 6)
            unread_emails = shared_mailbox.Items.Restrict(f"[Categories] = '{self.category}' AND [UnRead] = True")
            emails_to_process = [email for email in unread_emails]
            print(f"Found {len(emails_to_process)} emails to process under category '{self.category}' with 'UnRead' = True.")
            if save_emails:
                saved_emails = 0
                saved_attachments = 0
                not_saved_subjects = []
                incorrect_subjects = []
                for item in emails_to_process:
                    time.sleep(2)
                    try:
                        sender_email = item.SenderEmailAddress
                        sender_name_match = re.search(r'/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP.*?-([A-Za-z]+)', sender_email)
                        sender_name = sender_name_match.group(1) if sender_name_match else sender_email
                        print(f"Processing email from: {sender_name}")
                        if sender_name.lower() in [sender.lower() for sender in self.target_senders]:
                            if item.Attachments.Count == 0:
                                # Email has no attachments
                                emails_with_no_attachments.append(item.Subject)
                                print(f"Email with subject '{item.Subject}' has no attachments and will not be processed.")
                                continue  # Skip to the next email

                            # If we reach here, the email has attachments
                            subject_correct = True
                            saved_attachment_paths = []
                            has_pdf = False
                            has_nvf_new_vendor = False
                            save_msg_once = False
                            if item.Attachments.Count > 0:
                                for attachment in item.Attachments:
                                    if attachment.FileName.lower().endswith('.pdf'):
                                        has_pdf = True
                                    if 'nvf' in attachment.FileName.lower() or 'new vendor' in attachment.FileName.lower() or 'vendor' in attachment.FileName.lower():
                                        has_nvf_new_vendor = True
                                    if attachment.FileName.lower().endswith('.xlsx'):
                                        new_filename = self.extract_filename_from_subject(item.Subject)
                                        transformer = CharacterTransformer()
                                        new_filename = transformer.transform_to_swift_accepted_characters([new_filename])[0]
                                        new_filename = re.sub(r'[\/:*?"<>|\t]', ' ', new_filename)
                                        new_filename = re.sub(r'[^A-Za-z0-9\s\-\–;]', '', new_filename)
                                        if ';' not in new_filename and ';' not in item.Subject:
                                            print(f"Invalid filename generated from subject: {new_filename}")
                                            subject_correct = False
                                            incorrect_subjects.append(item.Subject)
                                            continue
                                        unique_attachment_filename = self.get_unique_filename(self.attachment_save_path, new_filename, '.xlsx')
                                        attachment_path = os.path.join(self.attachment_save_path, f"{unique_attachment_filename}.xlsx")
                                        attachment.SaveAsFile(attachment_path)
                                        print(f"Saved attachment: {attachment_path}")
                                        saved_attachment_paths.append(attachment_path)  # Track saved attachments
                                        saved_attachments += 1
                                        if not save_msg_once:
                                            unique_msg_filename = self.get_unique_filename(self.msg_save_path, new_filename, ' approval.msg')
                                            approval_msg_path = os.path.join(self.msg_save_path, f"{unique_msg_filename} approval.msg")
                                            item.SaveAs(approval_msg_path)
                                            print(f"Saved Outlook message: {approval_msg_path}")
                                            save_msg_once = True  # Ensure only one message file is saved
                                if has_pdf:
                                    self.emails_with_pdf.append(item.Subject)
                                if has_nvf_new_vendor:
                                    self.emails_with_nvf_new_vendor.append(item.Subject)
                                    print(f"NVF is present in the email with subject: {item.Subject}")
                                if subject_correct and not has_nvf_new_vendor:
                                    self.mark_email_as_read(item, mark_as_read)
                                    saved_emails += 1
                                    self.processed_emails[item.Subject] = saved_attachment_paths  # Track processed email
                                elif has_nvf_new_vendor:
                                    self.clean_up_files(saved_attachment_paths, approval_msg_path)
                    except Exception as e:
                        print(f"Error processing email from {sender_email}: {str(e)}")
                        not_saved_subjects.append(item.Subject)
                print(f"Total saved emails: {saved_emails}")
                actual_saved_attachments = self.count_files_in_directory(self.attachment_save_path)
                print(f"Total saved attachments: {actual_saved_attachments}")
                if not_saved_subjects:
                    print("Emails not saved due to errors:")
                    for subject in not_saved_subjects:
                        print(subject)
                if incorrect_subjects:
                    print("Emails with incorrect subject format:")
                    for subject in incorrect_subjects:
                        print(subject)
                # Additional block to copy attachments to the utils/pmt_run directory
                self.copy_attachments_to_pmt_run()

                # Create log file
                log_directory = r"C:\IT project3\utils\logs"
                os.makedirs(log_directory, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                log_filename = f"process_log_{timestamp}.txt"
                log_path = os.path.join(log_directory, log_filename)

                # Instantiate CharacterTransformer
                transformer = CharacterTransformer()

                with open(log_path, 'w', encoding='utf-8') as log_file:
                    log_file.write("Summary of processed emails and saved attachments:\n")
                    if self.processed_emails:
                        for subject, attachments in self.processed_emails.items():
                            # Transform the subject
                            transformed_subject = transformer.transform_to_swift_accepted_characters([subject])[0]
                            log_file.write(f"Email subject: {transformed_subject}\n")
                            # Transform the attachment names
                            transformed_attachments = transformer.transform_to_swift_accepted_characters(attachments)
                            for attachment in transformed_attachments:
                                log_file.write(f" - Saved attachment: {attachment}\n")

                    actual_saved_attachments = self.count_files_in_directory(self.attachment_save_path)
                    if saved_attachments != actual_saved_attachments:
                        log_file.write(f"Discrepancy detected: expected {saved_attachments} attachments, but found {actual_saved_attachments} in directory.\n")

                    if self.emails_with_pdf:
                        log_file.write("Emails with PDF attachments:\n")
                        for subject in self.emails_with_pdf:
                            log_file.write(f"{subject}\n")
                        print("PDF attachments are present in the following emails:")
                        for subject in self.emails_with_pdf:
                            print(subject)

                    if self.emails_with_nvf_new_vendor:
                        log_file.write("Emails with NVF or New Vendor in attachment filenames:\n")
                        for subject in self.emails_with_nvf_new_vendor:
                            log_file.write(f"{subject}\n")
                        print("NVF or New Vendor is present in the following emails:")
                        for subject in self.emails_with_nvf_new_vendor:
                            print(subject)

                    if emails_with_no_attachments:
                        log_file.write("Emails with no attachments:\n")
                        for subject in emails_with_no_attachments:
                            log_file.write(f"{subject}\n")
                        print("Emails with no attachments:")
                        for subject in emails_with_no_attachments:
                            print(subject)

                    if not_saved_subjects:
                        print("Emails not saved due to errors:")
                        for subject in not_saved_subjects:
                            print(subject)
                                    
                print(f"Log file created at {log_path}")

            else:
                print("Email saving is disabled. No emails were saved.")
        else:
            print(f"Could not resolve the recipient: {SHARED_MAILBOX_EMAIL}")


    def clean_up_files(self, attachment_paths, msg_path):
        for attachment_path in attachment_paths:
            try:
                os.remove(attachment_path)
                print(f"Deleted attachment: {attachment_path}")
            except Exception as e:
                print(f"Error deleting attachment {attachment_path}: {str(e)}")
        try:
            os.remove(msg_path)
            print(f"Deleted Outlook message: {msg_path}")
        except Exception as e:
            print(f"Error deleting Outlook message {msg_path}: {str(e)}")




    def copy_attachments_to_pmt_run(self):
        # Get the absolute path of the script
        script_path = os.path.dirname(os.path.abspath(__file__))
        # Define the pmt_run directory relative to the script location
        pmt_run_dir = os.path.join(script_path, "utils", "pmt_run")
        os.makedirs(pmt_run_dir, exist_ok=True)
        for filename in os.listdir(self.attachment_save_path):
            source_file = os.path.join(self.attachment_save_path, filename)
            destination_file = os.path.join(pmt_run_dir, filename)
            shutil.copy(source_file, destination_file)
        print(f"All files copied to {pmt_run_dir}")

    def extract_filename_from_subject(self, subject):
        match = re.search(r';\s*(.*)', subject)
        if match:
            return match.group(1)
        else:
            return subject

    def count_files_in_directory(self, directory):
        return len([name for name in os.listdir(directory) if os.path.isfile(os.path.join(directory, name))])

class ExcelTransposer:
    def __init__(self, filename: str):
        self.filename = filename
        self.workbook: Workbook = load_workbook(filename)
        self.sheet: Worksheet = self.workbook.active

    def set_active_sheet(self, sheet_name: str) -> None:
        if sheet_name in self.workbook.sheetnames:
            self.sheet = self.workbook[sheet_name]
        else:
            raise ValueError(f"Sheet '{sheet_name}' does not exist in the workbook.")

    def transpose_cells_to_table(self) -> None:
        try:
            data = []
            for row in self.sheet.iter_rows(values_only=True):
                data.append(row)
            transposed_data = list(map(list, zip(*data)))
            transposed_sheet = self.workbook.create_sheet(title="Transposed")
            for row_idx, row_data in enumerate(transposed_data):
                for col_idx, cell_value in enumerate(row_data):
                    column_letter = get_column_letter(col_idx + 1)
                    transposed_sheet[f"{column_letter}{row_idx + 1}"] = cell_value
            self.auto_adjust_column_width(transposed_sheet)
            #if the values in any row from columns F and G are both "None", print a warning with naming the row from value of the A column in that row
            for row in transposed_sheet.iter_rows(min_row=2, max_row=transposed_sheet.max_row, min_col=6, max_col=7):
                if row[0].value == "None" and row[1].value == "None":
                    print(f"Warning: There are missing bank account information in row {row[0].row}. The filename is {transposed_sheet[f'A{row[0].row}'].value}")
            #hide H column in the transposed sheet
            transposed_sheet.column_dimensions['H'].hidden = True            
            self.workbook.save(self.filename)
            print("Transposed data saved and column widths adjusted successfully.")
        except Exception as e:
            print(f"An error occurred while transposing the data: {e}")

    def auto_adjust_column_width(self, sheet: Worksheet) -> None:
        for column in sheet.columns:
            max_length = 0
            column_cells = [cell for cell in column]
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = adjusted_width

class KejsarProcessor:
    def __init__(self, base_dir: str):
        self.source_dir = os.path.join(base_dir, "utils", "pmt_run")
        self.error_dir = os.path.join(self.source_dir, "error_files")
        self.combined_file = os.path.join(base_dir, "utils", "combined_file.xlsx")
        self.case_list_folder = os.path.join(base_dir, "utils")
        self.comparison_file = os.path.join(base_dir, "utils", "comparison_file.xlsx")
        self.excel_files = list(Path(self.source_dir).glob("*.xlsx"))
        self.mismatched_cases = []
        os.makedirs(self.error_dir, exist_ok=True)
        self.comparison_data = self.load_comparison_data()

    def load_comparison_data(self):
        """Load data from comparison_file.xlsx with both cleaned and original values."""
        comparison_data = {}
        wb = load_workbook(filename=self.comparison_file)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Store cleaned key -> original bank name mapping
            if row[2] is not None:  # Column C value
                cleaned_key = str(row[2]).replace(" ", "").replace("-", "")
                comparison_data[cleaned_key] = row[0]  # Column A value (original bank name)
        return comparison_data

    def clean_value(self, value):
        """Minimal cleaning just for matching keys."""
        if value is not None:
            return str(value).replace(" ", "").replace("-", "")
        return value

    def match_bank_name(self, extra_cell_3_value, extra_cell_4_value):
        """Match cleaned values against comparison data and return original bank name."""
        if extra_cell_3_value is not None:
            extra_cell_3_value = str(extra_cell_3_value).replace(" ", "").replace("-", "")
        if extra_cell_4_value is not None:
            extra_cell_4_value = str(extra_cell_4_value).replace(" ", "").replace("-", "")

        # Look up original bank name using cleaned values
        bank_name = self.comparison_data.get(extra_cell_3_value) or self.comparison_data.get(extra_cell_4_value)
        return bank_name
    
    def process_files(self):
        i = 0
        while i < len(self.excel_files):
            excel_file = self.excel_files[i]
            try:
                wb = load_workbook(filename=excel_file)
                sheet_name = "Sheet1"
                if sheet_name not in wb.sheetnames:
                    print(f"Skipping and moving file {excel_file.name} as '{sheet_name}' worksheet does not exist.")
                    shutil.move(str(excel_file), os.path.join(self.error_dir, excel_file.name))
                    self.excel_files.pop(i)
                    continue
            except Exception as e:
                print(f"An error occurred with file {excel_file}: {e}")
                shutil.move(str(excel_file), os.path.join(self.error_dir, excel_file.name))
                self.excel_files.pop(i)
                continue
            i += 1

    def collect_values(self):
        pass #collect values from excel files cells


    def create_combined_excel(self, values_excel_files):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Sheet"
        
        header_list = ["uni name", "candidate name", "case nr", "amount", "currency", "acc number"]
        
        # Write headers and data
        for i, header in enumerate(header_list):
            worksheet[f"A{i+1}"] = header
        
        for i, excel_file in enumerate(values_excel_files):
            column_letter = get_column_letter(i + 2)
            worksheet[f"{column_letter}1"] = excel_file
            for j, value in enumerate(values_excel_files[excel_file]):
                if j == len(values_excel_files[excel_file]) - 3:
                    if value is not None and not isinstance(value, int):
                        value = str(value).replace("\n", "").replace("\xa0", "").replace("\t", "")
                worksheet[f"{column_letter}{j + 2}"] = value

    def adjust_columns(self):
        #transposer = ExcelTransposer(self.combined_file)
        #transposer.transpose_cells_to_table()
        wide = Wide(self.combined_file, "Transposed", self.source_dir)
        wide.auto_adjust_column_width()

    def process_case_list(self):
        case_list = CaseList(self.source_dir, self.case_list_folder)
        case_list.process_excel_files()

    def compare_files(self):
        comparator = ExcelComparator(self.combined_file, self.comparison_file, self.source_dir)
        comparator.compare_and_append()

    def run(self):
        self.process_files()
        values_excel_files = self.collect_values()
        self.create_combined_excel(values_excel_files)
        self.adjust_columns()
        self.process_case_list()
        self.compare_files()