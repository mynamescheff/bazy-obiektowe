import os
import win32com.client
from datetime import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
class OutlookProcessor:
    @staticmethod
    def check_unread_emails():
        """Checks for unread emails with Excel attachments without downloading them"""
        try:
            # Connect to Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # Inbox
            unread_emails = inbox.Items.Restrict("[Unread] = True")
            
            excel_email_count = 0
            
            for email in unread_emails:
                try:
                    if email.Attachments.Count > 0:
                        for attachment in email.Attachments:
                            if attachment.FileName.lower().endswith('.xlsx'):
                                excel_email_count += 1
                                break  # Count each email only once
                except Exception:
                    continue
                    
            return excel_email_count
        except Exception as e:
            logger.error(f"Error accessing Outlook: {str(e)}")
            return 0
    
    def download_xlsx_from_unread_emails(download_folder="./rozklasowany/outlook"):
        """
        Downloads .xlsx files from unread emails in Outlook
        
        Args:
            download_folder (str): Folder where files will be saved
        """
        try:
            # Create download folder if it doesn't exist
            if not os.path.exists(download_folder):
                os.makedirs(download_folder)
                logger.info(f"Created download folder: {download_folder}")
            
            # Connect to Outlook
            logger.info("Connecting to Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Get the inbox folder
            inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            logger.info("Connected to Outlook inbox")
            
            # Get unread emails
            unread_emails = inbox.Items.Restrict("[Unread] = True")
            logger.info(f"Found {unread_emails.Count} unread emails")
            
            downloaded_count = 0
            
            # Process each unread email
            for email in unread_emails:
                try:
                    logger.info(f"Processing email: {email.Subject}")
                    
                    # Check if email has attachments
                    if email.Attachments.Count > 0:
                        logger.info(f"  Found {email.Attachments.Count} attachments")
                        
                        # Process each attachment
                        for attachment in email.Attachments:
                            filename = attachment.FileName
                            
                            # Check if attachment is an Excel file
                            if filename.lower().endswith('.xlsx'):
                                # Create unique filename with timestamp to avoid conflicts
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                sender_name = email.SenderName.replace(" ", "_").replace("<", "").replace(">", "")
                                safe_filename = f"{timestamp}_{sender_name}_{filename}"
                                
                                # Remove any invalid characters for Windows filenames
                                invalid_chars = '<>:"/\\|?*'
                                for char in invalid_chars:
                                    safe_filename = safe_filename.replace(char, "_")
                                
                                filepath = os.path.join(download_folder, safe_filename)
                                
                                # Save the attachment
                                attachment.SaveAsFile(filepath)
                                downloaded_count += 1
                                
                                logger.info(f"  Downloaded: {safe_filename}")
                                logger.info(f"    From: {email.SenderName}")
                                logger.info(f"    Subject: {email.Subject}")
                                logger.info(f"    Received: {email.ReceivedTime}")
                                
                            else:
                                logger.info(f"  Skipping non-Excel file: {filename}")
                    else:
                        logger.info("  No attachments found")
                        
                except Exception as e:
                    logger.error(f"Error processing email '{email.Subject}': {str(e)}")
                    continue
            
            logger.info(f"Download complete! Downloaded {downloaded_count} Excel files to '{download_folder}'")
            return downloaded_count
            
        except Exception as e:
            logger.error(f"Error accessing Outlook: {str(e)}")
            return 0

    def mark_emails_as_read(mark_read=False):
        """
        Optional: Mark processed emails as read
        
        Args:
            mark_read (bool): Whether to mark emails as read after processing
        """
        if not mark_read:
            return
            
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            unread_emails = inbox.Items.Restrict("[Unread] = True")
            
            for email in unread_emails:
                if email.Attachments.Count > 0:
                    for attachment in email.Attachments:
                        if attachment.FileName.lower().endswith('.xlsx'):
                            email.UnRead = False
                            logger.info(f"Marked as read: {email.Subject}")
                            break
                            
        except Exception as e:
            logger.error(f"Error marking emails as read: {str(e)}")

    if __name__ == "__main__":
        print("Outlook Excel File Downloader")
        print("=" * 50)
        
        # Configuration
        DOWNLOAD_FOLDER = "downloaded_excel_files"  # Change this to your preferred folder
        MARK_AS_READ = False  # Set to True if you want to mark emails as read after downloading
        
        # Run the download process
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