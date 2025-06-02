import os
import win32com.client
from datetime import datetime

class OutlookProcessor:
    @staticmethod
    def download_xlsx_from_unread_emails(download_folder="./rozklasowany/outlook", output_callback=None):
        try:
            if output_callback is None:
                output_callback = print

            if not os.path.exists(download_folder):
                os.makedirs(download_folder)
                output_callback(f"Created download folder: {download_folder}")

            output_callback("Connecting to Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            inbox = namespace.GetDefaultFolder(6)
            output_callback("Connected to Outlook inbox")

            unread_emails = inbox.Items.Restrict("[Unread] = True")
            output_callback(f"Found {unread_emails.Count} unread emails")
            
            downloaded_count = 0

            for email in unread_emails:
                try:
                    output_callback(f"Processing email: {email.Subject}")

                    if email.Attachments.Count > 0:
                        output_callback(f"  Found {email.Attachments.Count} attachments")

                        for attachment in email.Attachments:
                            filename = attachment.FileName

                            if filename.lower().endswith('.xlsx'):
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                sender_name = email.SenderName.replace(" ", "_").replace("<", "").replace(">", "")
                                safe_filename = f"{timestamp}_{sender_name}_{filename}"

                                invalid_chars = '<>:"/\\|?*'
                                for char in invalid_chars:
                                    safe_filename = safe_filename.replace(char, "_")
                                
                                filepath = os.path.join(download_folder, safe_filename)

                                attachment.SaveAsFile(filepath)
                                downloaded_count += 1
                                
                                output_callback(f"  Downloaded: {safe_filename}")
                                output_callback(f"    From: {email.SenderName}")
                                output_callback(f"    Subject: {email.Subject}")
                                output_callback(f"    Received: {email.ReceivedTime}")
                                
                            else:
                                output_callback(f"  Skipping non-Excel file: {filename}")
                    else:
                        output_callback("  No attachments found")
                        
                except Exception as e:
                    output_callback(f"Error processing email '{email.Subject}': {str(e)}")
                    continue
            
            output_callback(f"Download complete! Downloaded {downloaded_count} Excel files to '{download_folder}'")
            return downloaded_count
            
        except Exception as e:
            output_callback(f"Error accessing Outlook: {str(e)}")
            return 0

    @staticmethod
    def mark_emails_as_read(mark_read=False, output_callback=None):
        if not mark_read:
            return
            
        try:
            if output_callback is None:
                output_callback = print
                
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            unread_emails = inbox.Items.Restrict("[Unread] = True")
            
            for email in unread_emails:
                if email.Attachments.Count > 0:
                    for attachment in email.Attachments:
                        if attachment.FileName.lower().endswith('.xlsx'):
                            email.UnRead = False
                            output_callback(f"Marked as read: {email.Subject}")
                            break
                            
        except Exception as e:
            output_callback(f"Error marking emails as read: {str(e)}")

    @staticmethod
    def check_unread_emails(output_callback=None):
        try:
            if output_callback is None:
                output_callback = print
                
            output_callback("Connecting to Outlook to check unread emails...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            unread_emails = inbox.Items.Restrict("[Unread] = True")
            
            excel_email_count = 0
            
            for email in unread_emails:
                try:
                    if email.Attachments.Count > 0:
                        for attachment in email.Attachments:
                            if attachment.FileName.lower().endswith('.xlsx'):
                                excel_email_count += 1
                                break
                except Exception:
                    continue
                    
            output_callback(f"Found {excel_email_count} unread emails with Excel attachments")
            return excel_email_count
        except Exception as e:
            output_callback(f"Error accessing Outlook: {str(e)}")
            return 0

    if __name__ == "__main__":
        print("Outlook Excel File Downloader")
        print("=" * 50)

        DOWNLOAD_FOLDER = "downloaded_excel_files"
        MARK_AS_READ = False

        try:
            count = download_xlsx_from_unread_emails(DOWNLOAD_FOLDER)
            
            if count > 0:
                print(f"\nSuccessfully downloaded {count} Excel files!")
                print(f"Files saved to: {os.path.abspath(DOWNLOAD_FOLDER)}")
                
                if MARK_AS_READ:
                    mark_emails_as_read(True)
                    print("Emails marked as read.")
            else:
                print("\nNo Excel files found in unread emails.")
                
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
        except Exception as e:
            print(f"\nAn error occurred: {str(e)}")
            print("Make sure Outlook is running and you have the required permissions.")
        
        input("\nPress Enter to exit...")