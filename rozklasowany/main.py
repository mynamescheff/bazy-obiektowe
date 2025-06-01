from PyQt6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, QCheckBox, QTextEdit,
                             QFileDialog, QMessageBox, QComboBox, QGroupBox)
from PyQt6.QtCore import QTimer
import threading
import sys
import os
from modules.excel_data_scraper import ExcelDataScraper
from modules.outlook_processor import OutlookProcessor
from modules.case_list import CaseList
from utils.database_handler import DatabaseHandler, COMBINED_DB_PATH_FOR_VERIFICATION, BANK_ACC_DB_PATH_FOR_VERIFICATION
from modules import relational_db_operations
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QIcon

class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Processor Tool")
        self.setGeometry(100, 100, 850, 700)

        # Initialize last directory
        self.last_directory = os.path.dirname(os.path.abspath(__file__))
        self.last_dir_path = ".\\rozklasowany\\last_dir"
        self.last_dir_file = os.path.join(self.last_dir_path, "last_directory.txt")

        # Create directory if it doesn't exist
        if not os.path.exists(self.last_dir_path):
            os.makedirs(self.last_dir_path)

        # Load last directory from file
        if os.path.exists(self.last_dir_file):
            with open(self.last_dir_file, "r") as f:
                self.last_directory = f.read().strip()

        # if the last directory is empty, set it to the current directory where the script is located
        if not self.last_directory:
            self.last_directory = os.path.dirname(os.path.abspath(__file__))

        # if the last directory is invalid, set it to the current directory where the script is located
        if not os.path.exists(self.last_directory):
            self.last_directory = os.path.dirname(os.path.abspath(__file__))

        # Status Bar
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

        # Initialize Components
        self.excel_scraper = ExcelDataScraper()
        self.db_handler = DatabaseHandler(status_var=None)  # Status handled via status_bar

        # Tab Widget
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        # Create Tabs
        self.outlook_tab = QWidget()
        self.case_list_tab = QWidget()
        self.scrape_tab = QWidget()
        self.db_utils_tab = QWidget()

        self.tab_widget.addTab(self.outlook_tab, "Outlook Processor")
        self.tab_widget.addTab(self.case_list_tab, "Case List")
        self.tab_widget.addTab(self.scrape_tab, "Excel Scraper")
        self.tab_widget.addTab(self.db_utils_tab, "Database Utilities")

        # Setup Tabs
        self.setup_outlook_tab()
        self.setup_case_list_tab()
        self.setup_scrape_tab()
        self.setup_db_utils_tab()

    ### Tab Setup Methods ###

    def setup_outlook_tab(self):
        layout = QVBoxLayout()

        # Form Layout for Inputs
        form_layout = QFormLayout()

        # Email Category
        self.category_entry = QLineEdit()
        self.category_entry.setText("Approval")
        form_layout.addRow("Email Category:", self.category_entry)

        # Target Senders
        self.senders_entry = QLineEdit()
        self.senders_entry.setText("Sender1,Sender2")
        form_layout.addRow("Target Senders (comma-separated):", self.senders_entry)

        # Attachments Save Path
        self.attachment_path_entry = QLineEdit()
        self.attachment_path_entry.setText("C:/Attachments")
        attachment_path_button = QPushButton("Browse")
        attachment_path_button.clicked.connect(lambda: self.browse_directory(self.attachment_path_entry))
        attachment_layout = QHBoxLayout()
        attachment_layout.addWidget(self.attachment_path_entry)
        attachment_layout.addWidget(attachment_path_button)
        form_layout.addRow("Attachments Save Path:", attachment_layout)

        # Messages Save Path
        self.msg_path_entry = QLineEdit()
        self.msg_path_entry.setText("C:/Messages")
        msg_path_button = QPushButton("Browse")
        msg_path_button.clicked.connect(lambda: self.browse_directory(self.msg_path_entry))
        msg_layout = QHBoxLayout()
        msg_layout.addWidget(self.msg_path_entry)
        msg_layout.addWidget(msg_path_button)
        form_layout.addRow("Messages Save Path:", msg_layout)

        # Checkboxes
        self.mark_as_read_check = QCheckBox("Mark emails as read")
        self.mark_as_read_check.setChecked(True)
        form_layout.addRow(self.mark_as_read_check)

        self.save_emails_check = QCheckBox("Save emails")
        self.save_emails_check.setChecked(True)
        form_layout.addRow(self.save_emails_check)

        layout.addLayout(form_layout)

        # Buttons
        buttons_layout = QHBoxLayout()
        check_unread_button = QPushButton("Check Unread Emails")
        check_unread_button.clicked.connect(self.check_unread_emails)
        process_emails_button = QPushButton("Process Emails")
        process_emails_button.clicked.connect(self.process_emails)
        buttons_layout.addWidget(check_unread_button)
        buttons_layout.addWidget(process_emails_button)
        layout.addLayout(buttons_layout)

        # Results Display
        layout.addWidget(QLabel("Processing Results:"))
        self.outlook_result_text = QTextEdit()
        self.outlook_result_text.setReadOnly(True)
        layout.addWidget(self.outlook_result_text)

        self.outlook_tab.setLayout(layout)

    def setup_case_list_tab(self):
        layout = QVBoxLayout()

        # Form Layout for Inputs
        form_layout = QFormLayout()

        # Excel Files Folder
        self.excel_folder_entry = QLineEdit()
        excel_folder_button = QPushButton("Browse")
        excel_folder_button.clicked.connect(lambda: self.browse_directory(self.excel_folder_entry))
        excel_folder_layout = QHBoxLayout()
        excel_folder_layout.addWidget(self.excel_folder_entry)
        excel_folder_layout.addWidget(excel_folder_button)
        form_layout.addRow("Excel Files Folder:", excel_folder_layout)

        # List Output Folder
        self.list_folder_entry = QLineEdit()
        list_folder_button = QPushButton("Browse")
        list_folder_button.clicked.connect(lambda: self.browse_directory(self.list_folder_entry))
        list_folder_layout = QHBoxLayout()
        list_folder_layout.addWidget(self.list_folder_entry)
        list_folder_layout.addWidget(list_folder_button)
        form_layout.addRow("List Output Folder:", list_folder_layout)

        layout.addLayout(form_layout)

        # Button
        process_case_list_button = QPushButton("Process Case List")
        process_case_list_button.clicked.connect(self.process_case_list)
        layout.addWidget(process_case_list_button)

        # Results Display
        layout.addWidget(QLabel("Processing Results:"))
        self.case_list_result_text = QTextEdit()
        self.case_list_result_text.setReadOnly(True)
        layout.addWidget(self.case_list_result_text)

        self.case_list_tab.setLayout(layout)

    def setup_scrape_tab(self):
        layout = QVBoxLayout()

        # Excel Files Directory
        self.scrape_dir_entry = QLineEdit()
        scrape_dir_button = QPushButton("Browse")
        scrape_dir_button.clicked.connect(lambda: self.browse_directory(self.scrape_dir_entry))
        scrape_dir_layout = QHBoxLayout()
        scrape_dir_layout.addWidget(self.scrape_dir_entry)
        scrape_dir_layout.addWidget(scrape_dir_button)
        layout.addWidget(QLabel("Excel Files Directory:"))
        layout.addLayout(scrape_dir_layout)

        # Range Selection
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("Cell Range:"))
        range_layout.addWidget(QLabel("From:"))
        self.range_start_entry = QLineEdit()
        self.range_start_entry.setText("A2")
        range_layout.addWidget(self.range_start_entry)
        range_layout.addWidget(QLabel("To:"))
        self.range_end_entry = QLineEdit()
        self.range_end_entry.setText("G2")
        range_layout.addWidget(self.range_end_entry)
        layout.addLayout(range_layout)

        # Read Headers Checkbox
        self.read_headers_check = QCheckBox("Read headers from first row")
        self.read_headers_check.setChecked(True)
        layout.addWidget(self.read_headers_check)

        # Buttons
        buttons_layout = QHBoxLayout()
        scrape_button = QPushButton("Scrape Excel Files")
        scrape_button.clicked.connect(self.scrape_excel_files)
        export_excel_button = QPushButton("Export to Excel")
        export_excel_button.clicked.connect(self.export_to_excel)
        export_csv_button = QPushButton("Export to CSV")
        export_csv_button.clicked.connect(self.export_to_csv)
        files_to_db_button = QPushButton("Files to DB (.txt/.xlsx)")
        files_to_db_button.clicked.connect(self.add_to_database)
        buttons_layout.addWidget(scrape_button)
        buttons_layout.addWidget(export_excel_button)
        buttons_layout.addWidget(export_csv_button)
        buttons_layout.addWidget(files_to_db_button)
        layout.addLayout(buttons_layout)

        # Results Display
        layout.addWidget(QLabel("Scraped Data:"))
        self.scrape_result_text = QTextEdit()
        self.scrape_result_text.setReadOnly(True)
        layout.addWidget(self.scrape_result_text)

        self.scrape_tab.setLayout(layout)

    def setup_db_utils_tab(self):
        layout = QVBoxLayout()

        # Bank Account Verification Section
        verify_group = QGroupBox("Bank Account Verification")
        verify_layout = QVBoxLayout()
        verify_desc = (f"Checks bank accounts from:\n'{os.path.abspath(COMBINED_DB_PATH_FOR_VERIFICATION)}'\n"
                       f"against:\n'{os.path.abspath(BANK_ACC_DB_PATH_FOR_VERIFICATION)}'.\n"
                       f"Ensure files exist and are formatted (table 'data', columns 'university', 'bank account').")
        verify_label = QLabel(verify_desc)
        verify_label.setWordWrap(True)
        verify_layout.addWidget(verify_label)
        verify_button = QPushButton("Verify Bank Accounts")
        verify_button.clicked.connect(self.run_verify_bank_accounts)
        verify_layout.addWidget(verify_button)
        verify_group.setLayout(verify_layout)
        layout.addWidget(verify_group)

        # Relational Database Project Section
        project_group = QGroupBox(f"Relational Project Database ({os.path.basename(relational_db_operations.PROJECT_DB_PATH)})")
        project_layout = QVBoxLayout()

        # Setup Buttons
        setup_layout = QHBoxLayout()
        setup_schema_button = QPushButton("Setup/Verify Project Schema")
        setup_schema_button.clicked.connect(self.run_setup_project_schema)
        populate_db_button = QPushButton("Populate DB from Combined.db")
        populate_db_button.clicked.connect(self.run_populate_project_data_from_combined_db)
        setup_layout.addWidget(setup_schema_button)
        setup_layout.addWidget(populate_db_button)
        project_layout.addLayout(setup_layout)

        # Query Section
        query_layout = QFormLayout()
        self.db_uni_name_combo = QComboBox()
        self.db_uni_name_combo.setEditable(False)
        query_layout.addRow("University Name:", self.db_uni_name_combo)
        refresh_uni_button = QPushButton("Refresh Uni List")
        refresh_uni_button.clicked.connect(self.load_university_combo_data)
        query_layout.addRow("", refresh_uni_button)
        show_cases_button = QPushButton("Show Cases for University")
        show_cases_button.clicked.connect(self.run_display_cases_for_university)
        query_layout.addRow("", show_cases_button)

        self.db_case_num_combo = QComboBox()
        self.db_case_num_combo.setEditable(False)
        query_layout.addRow("Case Number:", self.db_case_num_combo)
        refresh_case_button = QPushButton("Refresh Case List")
        refresh_case_button.clicked.connect(self.load_case_number_combo_data)
        query_layout.addRow("", refresh_case_button)
        show_uni_button = QPushButton("Show University for Case")
        show_uni_button.clicked.connect(self.run_display_university_for_case)
        query_layout.addRow("", show_uni_button)

        show_users_button = QPushButton("Show Users with Bank Accounts (All)")
        show_users_button.clicked.connect(self.run_display_users_with_bank_accounts)
        query_layout.addRow(show_users_button)

        project_layout.addLayout(query_layout)
        project_group.setLayout(project_layout)
        layout.addWidget(project_group)

        # Results Display
        layout.addWidget(QLabel("Output / Results (also check console):"))
        self.db_utils_result_text = QTextEdit()
        self.db_utils_result_text.setReadOnly(True)
        layout.addWidget(self.db_utils_result_text)

        self.db_utils_tab.setLayout(layout)

    ### Utility Methods ###

    def browse_directory(self, line_edit):
        # Use last_directory as the default directory for QFileDialog
        directory = QFileDialog.getExistingDirectory(self, "Select Directory", self.last_directory)
        if directory:
            line_edit.setText(directory)
            # Update last_directory and save to file
            self.last_directory = directory
            with open(self.last_dir_file, "w") as f:
                f.write(self.last_directory)

    def closeEvent(self, event):
        # Save last directory when closing the app
        with open(self.last_dir_file, "w") as f:
            f.write(self.last_directory)
        event.accept()

    def check_unread_emails(self):
        try:
            self.outlook_result_text.clear()
            category = self.category_entry.text()
            if not category:
                QMessageBox.critical(self, "Error", "Please enter an email category.")
                return
            processor = OutlookProcessor(category, [], "", "")
            unread_count = processor.list_unread_emails()
            self.outlook_result_text.append(f"Found {unread_count} unread emails with category '{category}'.")
            self.status_bar.showMessage(f"Found {unread_count} unread emails")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        except Exception as e:
            self.outlook_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error checking emails")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def process_emails(self):
        try:
            self.outlook_result_text.clear()
            category = self.category_entry.text()
            senders = [s.strip() for s in self.senders_entry.text().split(",") if s.strip()]
            attachment_path = self.attachment_path_entry.text()
            msg_path = self.msg_path_entry.text()

            if not all([category, senders, attachment_path, msg_path]):
                QMessageBox.critical(self, "Error", "Please fill in all required fields.")
                return

            self.outlook_result_text.append("Starting email processing...")
            self.status_bar.showMessage("Processing emails...")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

            def process_thread():
                processor = OutlookProcessor(category, senders, attachment_path, msg_path)
                processor.download_attachments_and_save_as_msg(
                    self.save_emails_check.isChecked(),
                    self.mark_as_read_check.isChecked()
                )
                QTimer.singleShot(0, lambda: self.update_outlook_results(processor))

            threading.Thread(target=process_thread, daemon=True).start()
        except Exception as e:
            self.outlook_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error processing emails")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def update_outlook_results(self, processor):
        self.outlook_result_text.append("Email processing completed.\n")
        self.outlook_result_text.append(f"Emails processed: {len(processor.processed_emails)}\n")
        if processor.emails_with_pdf:
            self.outlook_result_text.append("\nEmails with PDF attachments:\n")
            for subject in processor.emails_with_pdf:
                self.outlook_result_text.append(f"- {subject}\n")
        if processor.emails_with_nvf_new_vendor:
            self.outlook_result_text.append("\nEmails with NVF or New Vendor attachments:\n")
            for subject in processor.emails_with_nvf_new_vendor:
                self.outlook_result_text.append(f"- {subject}\n")
        self.status_bar.showMessage("Email processing completed")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def process_case_list(self):
        try:
            self.case_list_result_text.clear()
            excel_folder = self.excel_folder_entry.text()
            list_folder = self.list_folder_entry.text()

            if not all([excel_folder, list_folder]):
                QMessageBox.critical(self, "Error", "Please select both folders.")
                return

            self.case_list_result_text.append("Processing case list...")
            self.status_bar.showMessage("Processing case list...")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

            def process_thread():
                case_list = CaseList(excel_folder, list_folder)
                duplicate_counts, error_messages = case_list.process_excel_files(text_widget_update=self.case_list_result_text.append)
                # Capture the case list content
                case_list_content = self.get_case_list_content(list_folder)
                QTimer.singleShot(0, lambda: self.update_case_list_results(
                    duplicate_counts, error_messages, case_list_content))

            threading.Thread(target=process_thread, daemon=True).start()
        except Exception as e:
            self.case_list_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error processing case list")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def get_case_list_content(self, list_folder):
        """Read and return the content of case_list.txt"""
        list_file_path = os.path.join(list_folder, "case_list.txt")
        if os.path.exists(list_file_path):
            with open(list_file_path, "r", encoding="utf-8") as file:
                return file.read()
        return "No case list file found"

    def update_case_list_results(self, duplicate_counts, error_messages, case_list_content):
        self.case_list_result_text.append("Case list processing completed.\n")
        duplicates = sum(1 for count in duplicate_counts.values() if count > 0)
        self.case_list_result_text.append(f"Found {duplicates} duplicate cases.\n")
        
        if duplicates > 0:
            self.case_list_result_text.append("\nDuplicate cases:\n")
            for value, count in duplicate_counts.items():
                if count > 0:
                    self.case_list_result_text.append(f"- {value} (Duplicated {count} times)\n")
        
        if error_messages:
            self.case_list_result_text.append("\nErrors encountered:\n")
            for error in error_messages:
                self.case_list_result_text.append(f"- {error}\n")
        
        # Show case list content
        self.case_list_result_text.append("\nCase List Content:\n")
        self.case_list_result_text.append(case_list_content)
        
        self.status_bar.showMessage("Case list processing completed")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def update_case_list_results(self, duplicate_counts, error_messages):
        self.case_list_result_text.append("Case list processing completed.\n")
        duplicates = sum(1 for count in duplicate_counts.values() if count > 0)
        self.case_list_result_text.append(f"Found {duplicates} duplicate cases.\n")
        if duplicates > 0:
            self.case_list_result_text.append("\nDuplicate cases:\n")
            for value, count in duplicate_counts.items():
                if count > 0:
                    self.case_list_result_text.append(f"- {value} (Duplicated {count} times)\n")
        if error_messages:
            self.case_list_result_text.append("\nErrors encountered:\n")
            for error in error_messages:
                self.case_list_result_text.append(f"- {error}\n")
        self.status_bar.showMessage("Case list processing completed")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def scrape_excel_files(self):
        try:
            results, errors = [], []
            self.scrape_result_text.clear()
            directory = self.scrape_dir_entry.text()
            range_start = self.range_start_entry.text()
            range_end = self.range_end_entry.text()
            read_headers = self.read_headers_check.isChecked()

            if not directory:
                QMessageBox.critical(self, "Error", "Please select a directory with Excel files.")
                return

            self.scrape_result_text.append(f"Scraping Excel files in {directory}...")
            if not os.path.exists(directory):
                QMessageBox.critical(self, "Error", f"Directory '{directory}' does not exist.")
                return
            if not range_start or not range_end:
                QMessageBox.critical(self, "Error", "Please specify a valid cell range.")
                return
            if not range_start[0].isalpha() or not range_end[0].isalpha():
                QMessageBox.critical(self, "Error", "Cell range must start with a letter (e.g., A2).")
                return
            if not range_start[1:].isdigit() or not range_end[1:].isdigit():
                QMessageBox.critical(self, "Error", "Cell range must end with a number (e.g., A2).")
                return
            if range_start[0] > range_end[0]:
                QMessageBox.critical(self, "Error", "Start column must be before end column (e.g., A2 to G2).")
                return
            if int(range_start[1:]) > int(range_end[1:]):
                QMessageBox.critical(self, "Error", "Start row must be before end row (e.g., A2 to G2).")
                return
            self.scrape_result_text.append(f"Scraping range: {range_start} to {range_end} (Read Headers: {read_headers})")
            self.scrape_result_text.append("The files have been saved in the memory, please export them to CSV or Excel.")
            self.status_bar.showMessage("Scraping Excel files...")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

            self.excel_scraper.set_directory(directory)
            
            # Show a message box to let the user know scraping has ended
            QMessageBox.information(self, "Scraping Complete", "Scraping complete. You can now export the data to CSV or Excel.")

            def scrape_thread():
                # Get both results and errors
                results, errors = self.excel_scraper.scrape_excel_files(range_start, range_end, read_headers)
                QTimer.singleShot(0, lambda: self.update_scrape_results(results, errors))

            threading.Thread(target=scrape_thread, daemon=True).start()
        except Exception as e:
            self.scrape_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error scraping Excel files")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def update_scrape_results(self, results, errors):
        self.scrape_result_text.clear()
        
        # Show errors first
        if errors:
            self.scrape_result_text.append("Errors encountered:\n")
            for error in errors:
                self.scrape_result_text.append(f"- {error}\n")
            self.scrape_result_text.append("\n")
        
        self.scrape_result_text.append(f"Scraped {len(results)} Excel files.\n\n")
        
        headers = self.excel_scraper.get_headers()
        if headers:
            self.scrape_result_text.append("Headers found: " + ", ".join(headers) + "\n\n")
        
        if results:
            header_line = f"{'Filename':<30} | "
            if headers:
                for header in headers:
                    header_line += f"{str(header):<15} | "
            else:
                for key in results[0]["values"].keys():
                    header_line += f"{str(key):<15} | "
            self.scrape_result_text.append(header_line + "\n")
            self.scrape_result_text.append("-" * len(header_line) + "\n")
            for result in results:
                line = f"{result['filename']:<30} | "
                for key, value in result["values"].items():
                    if value is None:
                        value = ""
                    line += f"{str(value):<15} | "
                self.scrape_result_text.append(line + "\n")
        self.status_bar.showMessage(f"Scraped {len(results)} Excel files")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def export_to_excel(self):
        try:
            if not self.excel_scraper.get_results():
                QMessageBox.critical(self, "Error", "No data to export. Please scrape Excel files first.")
                return
            output_file, _ = QFileDialog.getSaveFileName(self, "Save Excel File", self.last_directory, "Excel Files (*.xlsx);;All Files (*)")
            if not output_file:
                return
            self.last_directory = os.path.dirname(output_file)
            success = self.excel_scraper.save_results_to_excel(output_file)
            if success:
                QMessageBox.information(self, "Success", f"Data exported to {output_file}")
                self.status_bar.showMessage("Data exported to Excel")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
    
            else:
                QMessageBox.critical(self, "Error", "Failed to export data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export error: {str(e)}")
            self.status_bar.showMessage("Error exporting data")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())


    def export_to_csv(self):
        try:
            if not self.excel_scraper.get_results():
                QMessageBox.critical(self, "Error", "No data to export. Please scrape Excel files first.")
                return
            output_file, _ = QFileDialog.getSaveFileName(self, "Save CSV File", self.last_directory, "CSV Files (*.csv);;All Files (*)")
            if not output_file:
                return
            self.last_directory = os.path.dirname(output_file)
            success = self.excel_scraper.save_results_to_csv(output_file)
            if success:
                QMessageBox.information(self, "Success", f"Data exported to {output_file}")
                self.status_bar.showMessage("Data exported to CSV")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
            else:
                QMessageBox.critical(self, "Error", "Failed to export data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export error: {str(e)}")
            self.status_bar.showMessage("Error exporting data")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def add_to_database(self):
        try:
            # Step 1: Select and convert .txt file
            txt_file_path, _ = QFileDialog.getOpenFileName(
                self, "Select a .txt File", self.last_directory, "Text Files (*.txt)"
            )
            if txt_file_path:
                self.last_directory = os.path.dirname(txt_file_path)
                self.status_bar.showMessage("Converting .txt file to database...")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
    
                try:
                    db_path = self.db_handler._convert_text_to_db(txt_file_path)
                    QMessageBox.information(
                        self, "Success",
                        f"Converted {os.path.basename(txt_file_path)} to {os.path.basename(db_path)}"
                    )
                except Exception as e:
                    QMessageBox.critical(
                        self, "Error",
                        f"Failed to convert {os.path.basename(txt_file_path)}: {str(e)}"
                    )
                    self.status_bar.showMessage("Error converting .txt file")
                    QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        
                    return
            else:
                self.status_bar.showMessage("TXT file selection canceled.")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
    
                return

            # Step 2: Select and convert .xlsx file
            excel_file_path, _ = QFileDialog.getOpenFileName(
                self, "Select an Excel File", self.last_directory, "Excel Files (*.xlsx *.xls)"
            )
            if excel_file_path:
                self.last_directory = os.path.dirname(excel_file_path)
                self.status_bar.showMessage("Converting Excel file to database...")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
    
                try:
                    db_path = self.db_handler._convert_excel_to_db(excel_file_path)
                    QMessageBox.information(
                        self, "Success",
                        f"Converted {os.path.basename(excel_file_path)} to {os.path.basename(db_path)}"
                    )
                except Exception as e:
                    QMessageBox.critical(
                        self, "Error",
                        f"Failed to convert {os.path.basename(excel_file_path)}: {str(e)}"
                    )
                    self.status_bar.showMessage("Error converting Excel file")
                    QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        
                    return
            else:
                self.status_bar.showMessage("Excel file selection canceled.")
                QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
    
                return

            self.status_bar.showMessage("File to DB conversion process complete.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Database error: {str(e)}")
            self.status_bar.showMessage("Error adding data to database")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())


    ### Database Utilities Methods ###

    def _append_db_util_result(self, message):
        self.db_utils_result_text.append(str(message))

    def run_verify_bank_accounts(self):
        self.db_utils_result_text.clear()
        self._append_db_util_result("Attempting to verify bank accounts...")
        self.status_bar.showMessage("Verifying bank accounts...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        self.db_handler.verify_bank_accounts_in_combined_db()
        self._append_db_util_result("Bank account verification process finished. Check status bar and pop-up messages.")

    def run_setup_project_schema(self):
        self.db_utils_result_text.clear()
        self._append_db_util_result(f"Attempting to set up/verify schema in '{relational_db_operations.PROJECT_DB_PATH}'...")
        self.status_bar.showMessage("Setting up project schema...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        try:
            relational_db_operations.setup_project_schema(relational_db_operations.PROJECT_DB_PATH)
            msg = f"Schema operation completed for {os.path.basename(relational_db_operations.PROJECT_DB_PATH)}."
            self._append_db_util_result(msg)
            QMessageBox.information(self, "Schema Setup", msg)
            self.status_bar.showMessage("Project schema setup complete.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

        except Exception as e:
            err_msg = f"Error during schema setup: {e}"
            self._append_db_util_result(err_msg)
            QMessageBox.critical(self, "Schema Error", err_msg)
            self.status_bar.showMessage("Error in schema setup.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())


    def run_populate_project_data_from_combined_db(self):
        self.db_utils_result_text.clear()
        self._append_db_util_result(f"Attempting to populate '{relational_db_operations.PROJECT_DB_PATH}' from '{os.path.basename(COMBINED_DB_PATH_FOR_VERIFICATION)}'...")
        self.status_bar.showMessage("Populating project DB from Combined.db...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        try:
            if not os.path.exists(relational_db_operations.PROJECT_DB_PATH):
                self._append_db_util_result(f"Database {relational_db_operations.PROJECT_DB_PATH} not found. Running schema setup first.")
                relational_db_operations.setup_project_schema(relational_db_operations.PROJECT_DB_PATH)
                self._append_db_util_result("Schema setup complete.")
            count = relational_db_operations.populate_project_data_from_combined_db(
                relational_db_operations.PROJECT_DB_PATH,
                COMBINED_DB_PATH_FOR_VERIFICATION,
                text_widget_update=self._append_db_util_result
            )
            msg = f"Population from Combined.db complete. Processed/updated {count} main entries."
            self._append_db_util_result(msg)
            QMessageBox.information(self, "Data Population", msg)
            self.status_bar.showMessage("Population from Combined.db complete.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

        except Exception as e:
            err_msg = f"Error during data population: {e}"
            self._append_db_util_result(err_msg)
            QMessageBox.critical(self, "Population Error", err_msg)
            self.status_bar.showMessage("Error in data population.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())


    def run_display_users_with_bank_accounts(self):
        self.db_utils_result_text.clear()
        self._append_db_util_result(f"Fetching users with bank accounts from '{relational_db_operations.PROJECT_DB_PATH}'...")
        self.status_bar.showMessage("Fetching users with accounts...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        try:
            df = relational_db_operations.display_users_with_bank_accounts(relational_db_operations.PROJECT_DB_PATH, text_widget_update=self._append_db_util_result)
            if df is not None and not df.empty:
                self._append_db_util_result("\n--- Users with Bank Accounts (via Cases) ---")
                self._append_db_util_result(df.to_string())
            elif df is not None and df.empty:
                self._append_db_util_result("No users found with bank accounts linked to their cases.")
            self.status_bar.showMessage("Displayed users with accounts.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        except Exception as e:
            self._append_db_util_result(f"Error displaying users: {e}")
            self.status_bar.showMessage("Error displaying users.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def run_display_cases_for_university(self):
        self.db_utils_result_text.clear()
        uni_name = self.db_uni_name_combo.currentText()
        if not uni_name:
            QMessageBox.critical(self, "Input Error", "Please select a University Name from the dropdown.")
            self.status_bar.showMessage("University not selected.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
            return
        self._append_db_util_result(f"Fetching cases for university '{uni_name}' from '{relational_db_operations.PROJECT_DB_PATH}'...")
        self.status_bar.showMessage(f"Fetching cases for {uni_name}...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        try:
            df = relational_db_operations.display_all_cases_for_university(relational_db_operations.PROJECT_DB_PATH, uni_name, text_widget_update=self._append_db_util_result)
            if df is not None and not df.empty:
                self._append_db_util_result(f"\n--- Cases for University: {uni_name} ---")
                self._append_db_util_result(df.to_string())
            elif df is not None and df.empty:
                self._append_db_util_result(f"No cases found for university: '{uni_name}'")
            self.status_bar.showMessage(f"Displayed cases for {uni_name}.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        except Exception as e:
            self._append_db_util_result(f"Error displaying cases: {e}")
            self.status_bar.showMessage("Error displaying cases.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def run_display_university_for_case(self):
        self.db_utils_result_text.clear()
        case_num = self.db_case_num_combo.currentText()
        if not case_num:
            QMessageBox.critical(self, "Input Error", "Please select a Case Number from the dropdown.")
            self.status_bar.showMessage("Case number not selected.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
            return
        self._append_db_util_result(f"Fetching university for case '{case_num}' from '{relational_db_operations.PROJECT_DB_PATH}'...")
        self.status_bar.showMessage(f"Fetching university for {case_num}...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        try:
            df = relational_db_operations.display_university_for_case(relational_db_operations.PROJECT_DB_PATH, case_num, text_widget_update=self._append_db_util_result)
            if df is not None and not df.empty:
                self._append_db_util_result(f"\n--- University for Case Number: {case_num} ---")
                self._append_db_util_result(df.to_string())
            elif df is not None and df.empty:
                self._append_db_util_result(f"No university found for case number: '{case_num}' (or case does not exist).")
            self.status_bar.showMessage(f"Displayed university for {case_num}.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        except Exception as e:
            self._append_db_util_result(f"Error displaying university: {e}")
            self.status_bar.showMessage("Error displaying university.")
            QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def load_university_combo_data(self):
        self.status_bar.showMessage("Loading university list...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        self._append_db_util_result("Fetching unique universities from bank_acc_db.db...")
        uni_list = relational_db_operations.get_unique_universities_from_bank_acc_db(
            BANK_ACC_DB_PATH_FOR_VERIFICATION,
            text_widget_update=self._append_db_util_result
        )
        self.db_uni_name_combo.clear()
        self.db_uni_name_combo.addItems(uni_list)
        if uni_list:
            self.db_uni_name_combo.setCurrentIndex(0)
            self.db_uni_name_combo.setEnabled(True)
        else:
            self.db_uni_name_combo.setEnabled(False)
        self.status_bar.showMessage("University list loaded.")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

    def load_case_number_combo_data(self):
        self.status_bar.showMessage("Loading case number list...")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())
        self._append_db_util_result(f"Fetching unique case numbers from {os.path.basename(relational_db_operations.CASE_LIST_DB_PATH)}...")
        case_list = relational_db_operations.get_unique_case_numbers_from_case_list_db(
            relational_db_operations.CASE_LIST_DB_PATH,
            text_widget_update=self._append_db_util_result
        )
        self.db_case_num_combo.clear()
        self.db_case_num_combo.addItems(case_list)
        if case_list:
            self.db_case_num_combo.setCurrentIndex(0)
            self.db_case_num_combo.setEnabled(True)
        else:
            self.db_case_num_combo.setEnabled(False)
        self.status_bar.showMessage("Case number list loaded.")
        QTimer.singleShot(2000, lambda: self.status_bar.clearMessage())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec())