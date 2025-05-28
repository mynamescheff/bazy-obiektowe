from PyQt6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QFormLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, QCheckBox, QTextEdit,
                             QFileDialog, QMessageBox)
from PyQt6.QtCore import QTimer
import threading
import sys
from rozklasowany.modules.outlook_processor import OutlookProcessor
from rozklasowany.modules.case_list import CaseList
from rozklasowany.modules.excel_data_scraper import ExcelDataScraper
from rozklasowany.utils.database_handler import DatabaseHandler

class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Processor Tool")
        self.setGeometry(100, 100, 800, 600)

        # Create tab widget
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        # Create tabs
        self.outlook_tab = QWidget()
        self.case_list_tab = QWidget()
        self.scrape_tab = QWidget()

        self.tab_widget.addTab(self.outlook_tab, "Outlook Processor")
        self.tab_widget.addTab(self.case_list_tab, "Case List")
        self.tab_widget.addTab(self.scrape_tab, "Excel Scraper")

        # Initialize components
        self.excel_scraper = ExcelDataScraper()

        # Setup tabs
        self.setup_outlook_tab()
        self.setup_case_list_tab()
        self.setup_scrape_tab()

        # Status bar
        self.status_bar = self.statusBar()
        self.status_bar.showMessage("Ready")

    def setup_outlook_tab(self):
        main_layout = QVBoxLayout()

        # Form layout for inputs
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
        self.attachment_path_button = QPushButton("Browse")
        self.attachment_path_button.clicked.connect(lambda: self.browse_directory(self.attachment_path_entry))
        attachment_path_layout = QHBoxLayout()
        attachment_path_layout.addWidget(self.attachment_path_entry)
        attachment_path_layout.addWidget(self.attachment_path_button)
        form_layout.addRow("Attachments Save Path:", attachment_path_layout)

        # Messages Save Path
        self.msg_path_entry = QLineEdit()
        self.msg_path_entry.setText("C:/Messages")
        self.msg_path_button = QPushButton("Browse")
        self.msg_path_button.clicked.connect(lambda: self.browse_directory(self.msg_path_entry))
        msg_path_layout = QHBoxLayout()
        msg_path_layout.addWidget(self.msg_path_entry)
        msg_path_layout.addWidget(self.msg_path_button)
        form_layout.addRow("Messages Save Path:", msg_path_layout)

        # Checkboxes
        self.mark_as_read_check = QCheckBox("Mark emails as read")
        self.mark_as_read_check.setChecked(True)
        form_layout.addRow(self.mark_as_read_check)

        self.save_emails_check = QCheckBox("Save emails")
        self.save_emails_check.setChecked(True)
        form_layout.addRow(self.save_emails_check)

        main_layout.addLayout(form_layout)

        # Buttons
        buttons_layout = QHBoxLayout()
        self.check_unread_button = QPushButton("Check Unread Emails")
        self.check_unread_button.clicked.connect(self.check_unread_emails)
        buttons_layout.addWidget(self.check_unread_button)

        self.process_emails_button = QPushButton("Process Emails")
        self.process_emails_button.clicked.connect(self.process_emails)
        buttons_layout.addWidget(self.process_emails_button)
        main_layout.addLayout(buttons_layout)

        # Results display
        self.outlook_result_label = QLabel("Processing Results:")
        main_layout.addWidget(self.outlook_result_label)
        self.outlook_result_text = QTextEdit()
        self.outlook_result_text.setReadOnly(True)
        main_layout.addWidget(self.outlook_result_text)

        self.outlook_tab.setLayout(main_layout)

    def setup_case_list_tab(self):
        main_layout = QVBoxLayout()

        # Form layout for inputs
        form_layout = QFormLayout()

        # Excel Files Folder
        self.excel_folder_entry = QLineEdit()
        self.excel_folder_button = QPushButton("Browse")
        self.excel_folder_button.clicked.connect(lambda: self.browse_directory(self.excel_folder_entry))
        excel_folder_layout = QHBoxLayout()
        excel_folder_layout.addWidget(self.excel_folder_entry)
        excel_folder_layout.addWidget(self.excel_folder_button)
        form_layout.addRow("Excel Files Folder:", excel_folder_layout)

        # List Output Folder
        self.list_folder_entry = QLineEdit()
        self.list_folder_button = QPushButton("Browse")
        self.list_folder_button.clicked.connect(lambda: self.browse_directory(self.list_folder_entry))
        list_folder_layout = QHBoxLayout()
        list_folder_layout.addWidget(self.list_folder_entry)
        list_folder_layout.addWidget(self.list_folder_button)
        form_layout.addRow("List Output Folder:", list_folder_layout)

        main_layout.addLayout(form_layout)

        # Button
        self.process_case_list_button = QPushButton("Process Case List")
        self.process_case_list_button.clicked.connect(self.process_case_list)
        main_layout.addWidget(self.process_case_list_button)

        # Results display
        self.case_list_result_label = QLabel("Processing Results:")
        main_layout.addWidget(self.case_list_result_label)
        self.case_list_result_text = QTextEdit()
        self.case_list_result_text.setReadOnly(True)
        main_layout.addWidget(self.case_list_result_text)

        self.case_list_tab.setLayout(main_layout)

    def setup_scrape_tab(self):
        main_layout = QVBoxLayout()

        # Excel Files Directory
        self.scrape_dir_entry = QLineEdit()
        self.scrape_dir_button = QPushButton("Browse")
        self.scrape_dir_button.clicked.connect(lambda: self.browse_directory(self.scrape_dir_entry))
        self.export_excel_button = QPushButton("Export to Excel")
        self.export_excel_button.clicked.connect(self.export_to_excel)
        scrape_dir_layout = QHBoxLayout()
        scrape_dir_layout.addWidget(self.scrape_dir_entry)
        scrape_dir_layout.addWidget(self.scrape_dir_button)
        scrape_dir_layout.addWidget(self.export_excel_button)
        main_layout.addWidget(QLabel("Excel Files Directory:"))
        main_layout.addLayout(scrape_dir_layout)

        # Range selection
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("Cell Range:"))
        range_layout.addWidget(QLabel("From:"))
        self.range_start_entry = QLineEdit()
        self.range_start_entry.setText("A2")
        range_layout.addWidget(self.range_start_entry)
        range_layout.addWidget(QLabel("To:"))
        self.range_end_entry = QLineEdit()
        self.range_end_entry.setText("F2")
        range_layout.addWidget(self.range_end_entry)
        main_layout.addLayout(range_layout)

        # Read headers checkbox
        self.read_headers_check = QCheckBox("Read headers from first row")
        self.read_headers_check.setChecked(True)
        main_layout.addWidget(self.read_headers_check)

        # Buttons
        buttons_layout = QHBoxLayout()
        self.scrape_button = QPushButton("Scrape Excel Files")
        self.scrape_button.clicked.connect(self.scrape_excel_files)
        buttons_layout.addWidget(self.scrape_button)
        self.export_csv_button = QPushButton("Export to CSV")
        self.export_csv_button.clicked.connect(self.export_to_csv)
        buttons_layout.addWidget(self.export_csv_button)
        self.add_to_db_button = QPushButton("Add to Database")
        self.add_to_db_button.clicked.connect(self.add_to_database)
        buttons_layout.addWidget(self.add_to_db_button)
        main_layout.addLayout(buttons_layout)

        # Results display
        main_layout.addWidget(QLabel("Scraped Data:"))
        self.scrape_result_text = QTextEdit()
        self.scrape_result_text.setReadOnly(True)
        self.scrape_result_text.setLineWrapMode(QTextEdit.NoWrap)
        main_layout.addWidget(self.scrape_result_text)

        self.scrape_tab.setLayout(main_layout)

    def browse_directory(self, line_edit):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            line_edit.setText(directory)

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
        except Exception as e:
            self.outlook_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error checking emails")

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
            processor = OutlookProcessor(category, senders, attachment_path, msg_path)

            def process_thread():
                processor.download_attachments_and_save_as_msg(
                    self.save_emails_check.isChecked(),
                    self.mark_as_read_check.isChecked()
                )
                QTimer.singleShot(0, lambda: self.update_outlook_results(processor))

            threading.Thread(target=process_thread, daemon=True).start()
        except Exception as e:
            self.outlook_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error processing emails")

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
            case_list = CaseList(excel_folder, list_folder)

            def process_thread():
                duplicate_counts, error_messages = case_list.process_excel_files()
                QTimer.singleShot(0, lambda: self.update_case_list_results(duplicate_counts, error_messages))

            threading.Thread(target=process_thread, daemon=True).start()
        except Exception as e:
            self.case_list_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error processing case list")

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

    def scrape_excel_files(self):
        try:
            self.scrape_result_text.clear()
            directory = self.scrape_dir_entry.text()
            range_start = self.range_start_entry.text()
            range_end = self.range_end_entry.text()
            read_headers = self.read_headers_check.isChecked()

            if not directory:
                QMessageBox.critical(self, "Error", "Please select a directory with Excel files.")
                return

            self.scrape_result_text.append(f"Scraping Excel files in {directory}...")
            self.status_bar.showMessage("Scraping Excel files...")
            self.excel_scraper.set_directory(directory)

            def scrape_thread():
                results = self.excel_scraper.scrape_excel_files(range_start, range_end, read_headers)
                QTimer.singleShot(0, lambda: self.update_scrape_results(results))

            threading.Thread(target=scrape_thread, daemon=True).start()
        except Exception as e:
            self.scrape_result_text.append(f"Error: {str(e)}")
            self.status_bar.showMessage("Error scraping Excel files")

    def update_scrape_results(self, results):
        self.scrape_result_text.clear()
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

    def export_to_excel(self):
        try:
            if not self.excel_scraper.get_results():
                QMessageBox.critical(self, "Error", "No data to export. Please scrape Excel files first.")
                return
            output_file, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
            if not output_file:
                return
            success = self.excel_scraper.save_results_to_excel(output_file)
            if success:
                QMessageBox.information(self, "Success", f"Data exported to {output_file}")
                self.status_bar.showMessage("Data exported to Excel")
            else:
                QMessageBox.critical(self, "Error", "Failed to export data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export error: {str(e)}")
            self.status_bar.showMessage("Error exporting data")

    def export_to_csv(self):
        try:
            if not self.excel_scraper.get_results():
                QMessageBox.critical(self, "Error", "No data to export. Please scrape Excel files first.")
                return
            output_file, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv);;All Files (*)")
            if not output_file:
                return
            success = self.excel_scraper.save_results_to_csv(output_file)
            if success:
                QMessageBox.information(self, "Success", f"Data exported to {output_file}")
                self.status_bar.showMessage("Data exported to CSV")
            else:
                QMessageBox.critical(self, "Error", "Failed to export data")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export error: {str(e)}")
            self.status_bar.showMessage("Error exporting data")

    def add_to_database(self):
        try:
            if not self.excel_scraper.get_results():
                QMessageBox.critical(self, "Error", "No data to add. Please scrape Excel files first.")
                return
            db_handler = DatabaseHandler()
            success = db_handler.insert_data(self.excel_scraper.get_results())
            if success:
                QMessageBox.information(self, "Success", "Data added to database successfully.")
                self.status_bar.showMessage("Data added to database")
            else:
                QMessageBox.critical(self, "Error", "Failed to add data to database")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Database error: {str(e)}")
            self.status_bar.showMessage("Error adding data to database")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec())