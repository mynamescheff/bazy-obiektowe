import win32com.client
import os
import re
import time
import shutil
from datetime import datetime
from ..utils.constants import SHARED_MAILBOX_EMAIL

SHARED_MAILBOX_EMAIL = SHARED_MAILBOX_EMAIL  # Ensure this constant is defined in your constants module

class OutlookProcessor:
    def __init__(self, category, target_senders, attachment_save_path, msg_save_path):
        self.category = category
        self.target_senders = target_senders
        self.attachment_save_path = attachment_save_path
        self.msg_save_path = msg_save_path
        self.processed_emails = {}
        self.emails_with_pdf = []
        self.emails_with_nvf_new_vendor = []
        
        # Create directories if they don't exist
        os.makedirs(self.attachment_save_path, exist_ok=True)
        os.makedirs(self.msg_save_path, exist_ok=True)

    def initialize_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            return True
        except Exception as e:
            print(f"Error initializing Outlook: {str(e)}")
            return False

    def list_unread_emails(self):
        if not hasattr(self, 'outlook') or not hasattr(self, 'namespace'):
            if not self.initialize_outlook():
                return 0
                
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

    def transform_to_swift_accepted_characters(self, text_list):
        """Basic character transformation - simplified from what was referenced"""
        if not text_list:
            return []
            
        result = []
        for text in text_list:
            if text:
                # Remove potentially problematic characters
                clean_text = re.sub(r'[\/:*?"<>|\t]', ' ', str(text))
                result.append(clean_text)
            else:
                result.append("")
        return result

    def download_attachments_and_save_as_msg(self, save_emails, mark_as_read):
        if not hasattr(self, 'outlook') or not hasattr(self, 'namespace'):
            if not self.initialize_outlook():
                print("Failed to initialize Outlook. Cannot process emails.")
                return
                
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

                            # Process email with attachments
                            subject_correct = True
                            saved_attachment_paths = []
                            has_pdf = False
                            has_nvf_new_vendor = False
                            save_msg_once = False
                            approval_msg_path = ""
                            
                            if item.Attachments.Count > 0:
                                for attachment in item.Attachments:
                                    if attachment.FileName.lower().endswith('.pdf'):
                                        has_pdf = True
                                    if 'nvf' in attachment.FileName.lower() or 'new vendor' in attachment.FileName.lower() or 'vendor' in attachment.FileName.lower():
                                        has_nvf_new_vendor = True
                                    if attachment.FileName.lower().endswith('.xlsx'):
                                        new_filename = self.extract_filename_from_subject(item.Subject)
                                        new_filename = self.transform_to_swift_accepted_characters([new_filename])[0]
                                        
                                        if ';' not in new_filename and ';' not in item.Subject:
                                            print(f"Invalid filename generated from subject: {new_filename}")
                                            subject_correct = False
                                            incorrect_subjects.append(item.Subject)
                                            continue
                                            
                                        unique_attachment_filename = self.get_unique_filename(self.attachment_save_path, new_filename, '.xlsx')
                                        attachment_path = os.path.join(self.attachment_save_path, f"{unique_attachment_filename}.xlsx")
                                        attachment.SaveAsFile(attachment_path)
                                        print(f"Saved attachment: {attachment_path}")
                                        saved_attachment_paths.append(attachment_path)
                                        saved_attachments += 1
                                        
                                        if not save_msg_once:
                                            unique_msg_filename = self.get_unique_filename(self.msg_save_path, new_filename, ' approval.msg')
                                            approval_msg_path = os.path.join(self.msg_save_path, f"{unique_msg_filename} approval.msg")
                                            item.SaveAs(approval_msg_path)
                                            print(f"Saved Outlook message: {approval_msg_path}")
                                            save_msg_once = True
                                            
                                if has_pdf:
                                    self.emails_with_pdf.append(item.Subject)
                                if has_nvf_new_vendor:
                                    self.emails_with_nvf_new_vendor.append(item.Subject)
                                    print(f"NVF is present in the email with subject: {item.Subject}")
                                    
                                if subject_correct and not has_nvf_new_vendor:
                                    self.mark_email_as_read(item, mark_as_read)
                                    saved_emails += 1
                                    self.processed_emails[item.Subject] = saved_attachment_paths
                                elif has_nvf_new_vendor and approval_msg_path:
                                    self.clean_up_files(saved_attachment_paths, approval_msg_path)
                                    
                    except Exception as e:
                        print(f"Error processing email: {str(e)}")
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
                        
                # Copy attachments to the utils/pmt_run directory
                self.copy_attachments_to_pmt_run()

                # Create log file
                self.create_log_file(emails_with_no_attachments, not_saved_subjects)

            else:
                print("Email saving is disabled. No emails were saved.")
        else:
            print(f"Could not resolve the recipient: {SHARED_MAILBOX_EMAIL}")

    def create_log_file(self, emails_with_no_attachments, not_saved_subjects):
        log_directory = r"C:\IT project3\utils\logs"
        os.makedirs(log_directory, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"process_log_{timestamp}.txt"
        log_path = os.path.join(log_directory, log_filename)

        with open(log_path, 'w', encoding='utf-8') as log_file:
            log_file.write("Summary of processed emails and saved attachments:\n")
            if self.processed_emails:
                for subject, attachments in self.processed_emails.items():
                    # Transform the subject
                    transformed_subject = self.transform_to_swift_accepted_characters([subject])[0]
                    log_file.write(f"Email subject: {transformed_subject}\n")
                    # Transform the attachment names
                    transformed_attachments = self.transform_to_swift_accepted_characters(attachments)
                    for attachment in transformed_attachments:
                        log_file.write(f" - Saved attachment: {attachment}\n")

            actual_saved_attachments = self.count_files_in_directory(self.attachment_save_path)
            expected_attachments = sum(len(paths) for paths in self.processed_emails.values())
            if expected_attachments != actual_saved_attachments:
                log_file.write(f"Discrepancy detected: expected {expected_attachments} attachments, but found {actual_saved_attachments} in directory.\n")

            if self.emails_with_pdf:
                log_file.write("Emails with PDF attachments:\n")
                for subject in self.emails_with_pdf:
                    log_file.write(f"{subject}\n")

            if self.emails_with_nvf_new_vendor:
                log_file.write("Emails with NVF or New Vendor in attachment filenames:\n")
                for subject in self.emails_with_nvf_new_vendor:
                    log_file.write(f"{subject}\n")

            if emails_with_no_attachments:
                log_file.write("Emails with no attachments:\n")
                for subject in emails_with_no_attachments:
                    log_file.write(f"{subject}\n")

            if not_saved_subjects:
                log_file.write("Emails not saved due to errors:\n")
                for subject in not_saved_subjects:
                    log_file.write(f"{subject}\n")
                    
        print(f"Log file created at {log_path}")

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