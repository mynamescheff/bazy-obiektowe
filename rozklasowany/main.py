import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import sqlite3
from rozklasowany.modules.excel_data_scraper import ExcelDataScraper
from modules.excel_transposer import ExcelTransposer
from modules.outlook_processor import OutlookProcessor
from modules.case_list import CaseList
from utils.database_handler import DatabaseHandler

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor Tool")
        self.root.geometry("800x600")
        
        # Create a notebook with tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create frames for each tab
        self.outlook_tab = ttk.Frame(self.notebook)
        self.case_list_tab = ttk.Frame(self.notebook)
        self.transpose_tab = ttk.Frame(self.notebook)
        self.scrape_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.outlook_tab, text="Outlook Processor")
        self.notebook.add(self.case_list_tab, text="Case List")
        #self.notebook.add(self.transpose_tab, text="Excel Transpose")
        self.notebook.add(self.scrape_tab, text="Excel Scraper")
        
        # Setup tabs
        self.setup_outlook_tab()
        self.setup_case_list_tab()
        self.setup_transpose_tab()
        self.setup_scrape_tab()
        
        # Initialize components
        self.excel_scraper = ExcelDataScraper()
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        
        # Status bar at the bottom
        status_frame = ttk.Frame(root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)

    def export_to_excel(self):
        try:
            if not self.excel_scraper.get_results():
                messagebox.showerror("Error", "No data to export. Please scrape Excel files first.")
                return
                
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
            )
            
            if not output_file:
                return  # User cancelled
                
            success = self.excel_scraper.save_results_to_excel(output_file)
            
            if success:
                messagebox.showinfo("Success", f"Data exported to {output_file}")
                self.status_var.set("Data exported to Excel")
            else:
                messagebox.showerror("Error", "Failed to export data")
                
        except Exception as e:
            messagebox.showerror("Error", f"Export error: {str(e)}")
            self.status_var.set("Error exporting data")

    def setup_outlook_tab(self):
        # Create form for Outlook processor
        ttk.Label(self.outlook_tab, text="Email Category:").grid(row=0, column=0, sticky=W, padx=10, pady=5)
        self.category_entry = ttk.Entry(self.outlook_tab, width=40)
        self.category_entry.grid(row=0, column=1, sticky=W, padx=10, pady=5)
        self.category_entry.insert(0, "Approval")
        
        ttk.Label(self.outlook_tab, text="Target Senders (comma-separated):").grid(row=1, column=0, sticky=W, padx=10, pady=5)
        self.senders_entry = ttk.Entry(self.outlook_tab, width=40)
        self.senders_entry.grid(row=1, column=1, sticky=W, padx=10, pady=5)
        self.senders_entry.insert(0, "Sender1,Sender2")
        
        ttk.Label(self.outlook_tab, text="Attachments Save Path:").grid(row=2, column=0, sticky=W, padx=10, pady=5)
        self.attachment_path_frame = ttk.Frame(self.outlook_tab)
        self.attachment_path_frame.grid(row=2, column=1, sticky=W, padx=10, pady=5)
        self.attachment_path_entry = ttk.Entry(self.attachment_path_frame, width=30)
        self.attachment_path_entry.pack(side=tk.LEFT)
        self.attachment_path_entry.insert(0, "C:/Attachments")
        ttk.Button(self.attachment_path_frame, text="Browse", command=lambda: self.browse_directory(self.attachment_path_entry)).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.outlook_tab, text="Messages Save Path:").grid(row=3, column=0, sticky=W, padx=10, pady=5)
        self.msg_path_frame = ttk.Frame(self.outlook_tab)
        self.msg_path_frame.grid(row=3, column=1, sticky=W, padx=10, pady=5)
        self.msg_path_entry = ttk.Entry(self.msg_path_frame, width=30)
        self.msg_path_entry.pack(side=tk.LEFT)
        self.msg_path_entry.insert(0, "C:/Messages")
        ttk.Button(self.msg_path_frame, text="Browse", command=lambda: self.browse_directory(self.msg_path_entry)).pack(side=tk.LEFT, padx=5)
        
        self.mark_as_read_var = BooleanVar(value=True)
        ttk.Checkbutton(self.outlook_tab, text="Mark emails as read", variable=self.mark_as_read_var).grid(row=4, column=0, sticky=W, padx=10, pady=5)
        
        self.save_emails_var = BooleanVar(value=True)
        ttk.Checkbutton(self.outlook_tab, text="Save emails", variable=self.save_emails_var).grid(row=4, column=1, sticky=W, padx=10, pady=5)
        
        ttk.Button(self.outlook_tab, text="Check Unread Emails", command=self.check_unread_emails).grid(row=5, column=0, sticky=W, padx=10, pady=5)
        ttk.Button(self.outlook_tab, text="Process Emails", command=self.process_emails).grid(row=5, column=1, sticky=W, padx=10, pady=5)
        
        # Results display
        ttk.Label(self.outlook_tab, text="Processing Results:").grid(row=6, column=0, sticky=W, padx=10, pady=5)
        self.outlook_result_text = Text(self.outlook_tab, height=15, width=80)
        self.outlook_result_text.grid(row=7, column=0, columnspan=2, padx=10, pady=5)
        
        # Add scrollbar
        outlook_scrollbar = ttk.Scrollbar(self.outlook_tab, command=self.outlook_result_text.yview)
        outlook_scrollbar.grid(row=7, column=2, sticky='nsew')
        self.outlook_result_text.config(yscrollcommand=outlook_scrollbar.set)


    def setup_case_list_tab(self):
        ttk.Label(self.case_list_tab, text="Excel Files Folder:").grid(row=0, column=0, sticky=W, padx=10, pady=5)
        self.excel_folder_frame = ttk.Frame(self.case_list_tab)
        self.excel_folder_frame.grid(row=0, column=1, sticky=W, padx=10, pady=5)
        self.excel_folder_entry = ttk.Entry(self.excel_folder_frame, width=30)
        self.excel_folder_entry.pack(side=tk.LEFT)
        ttk.Button(self.excel_folder_frame, text="Browse", command=lambda: self.browse_directory(self.excel_folder_entry)).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.case_list_tab, text="List Output Folder:").grid(row=1, column=0, sticky=W, padx=10, pady=5)
        self.list_folder_frame = ttk.Frame(self.case_list_tab)
        self.list_folder_frame.grid(row=1, column=1, sticky=W, padx=10, pady=5)
        self.list_folder_entry = ttk.Entry(self.list_folder_frame, width=30)
        self.list_folder_entry.pack(side=tk.LEFT)
        ttk.Button(self.list_folder_frame, text="Browse", command=lambda: self.browse_directory(self.list_folder_entry)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(self.case_list_tab, text="Process Case List", command=self.process_case_list).grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        
        # Results display
        ttk.Label(self.case_list_tab, text="Processing Results:").grid(row=3, column=0, sticky=W, padx=10, pady=5)
        self.case_list_result_text = Text(self.case_list_tab, height=15, width=80)
        self.case_list_result_text.grid(row=4, column=0, columnspan=2, padx=10, pady=5)
        
        # Add scrollbar
        case_list_scrollbar = ttk.Scrollbar(self.case_list_tab, command=self.case_list_result_text.yview)
        case_list_scrollbar.grid(row=4, column=2, sticky='nsew')
        self.case_list_result_text.config(yscrollcommand=case_list_scrollbar.set)

    def setup_transpose_tab(self):
        ttk.Label(self.transpose_tab, text="Excel File:").grid(row=0, column=0, sticky=W, padx=10, pady=5)
        self.excel_file_frame = ttk.Frame(self.transpose_tab)
        self.excel_file_frame.grid(row=0, column=1, sticky=W, padx=10, pady=5)
        self.excel_file_entry = ttk.Entry(self.excel_file_frame, width=30)
        self.excel_file_entry.pack(side=tk.LEFT)
        ttk.Button(self.excel_file_frame, text="Browse", command=lambda: self.browse_file(self.excel_file_entry, [("Excel Files", "*.xlsx")])).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.transpose_tab, text="Sheet Name (optional):").grid(row=1, column=0, sticky=W, padx=10, pady=5)
        self.sheet_name_entry = ttk.Entry(self.transpose_tab, width=30)
        self.sheet_name_entry.grid(row=1, column=1, sticky=W, padx=10, pady=5)
        
        ttk.Button(self.transpose_tab, text="Transpose Excel", command=self.transpose_excel).grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        
        # Results display
        ttk.Label(self.transpose_tab, text="Transpose Results:").grid(row=3, column=0, sticky=W, padx=10, pady=5)
        self.transpose_result_text = Text(self.transpose_tab, height=15, width=80)
        self.transpose_result_text.grid(row=4, column=0, columnspan=2, padx=10, pady=5)
    
    def setup_scrape_tab(self):
        ttk.Label(self.scrape_tab, text="Excel Files Directory:").grid(row=0, column=0, sticky=W, padx=10, pady=5)
        self.scrape_dir_frame = ttk.Frame(self.scrape_tab)
        self.scrape_dir_frame.grid(row=0, column=1, sticky=W, padx=10, pady=5)
        self.scrape_dir_entry = ttk.Entry(self.scrape_dir_frame, width=30)
        self.scrape_dir_entry.pack(side=tk.LEFT)
        ttk.Button(self.scrape_dir_frame, text="Browse", command=lambda: self.browse_directory(self.scrape_dir_entry)).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.scrape_dir_frame, text="Export to Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        
        # Range selection
        range_frame = ttk.Frame(self.scrape_tab)
        range_frame.grid(row=1, column=0, columnspan=2, sticky=W, padx=10, pady=5)
        
        ttk.Label(range_frame, text="Cell Range:").pack(side=tk.LEFT, padx=5)
        ttk.Label(range_frame, text="From:").pack(side=tk.LEFT, padx=5)
        self.range_start_entry = ttk.Entry(range_frame, width=8)
        self.range_start_entry.pack(side=tk.LEFT, padx=5)
        self.range_start_entry.insert(0, "A2")
        ttk.Label(range_frame, text="To:").pack(side=tk.LEFT, padx=5)
        self.range_end_entry = ttk.Entry(range_frame, width=8)
        self.range_end_entry.pack(side=tk.LEFT, padx=5)
        self.range_end_entry.insert(0, "F2")
        
        # Read headers option
        self.read_headers_var = BooleanVar(value=True)
        ttk.Checkbutton(self.scrape_tab, text="Read headers from first row", variable=self.read_headers_var).grid(row=2, column=0, columnspan=2, sticky=W, padx=10, pady=5)
        
        # Buttons
        buttons_frame = ttk.Frame(self.scrape_tab)
        buttons_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Scrape Excel Files", command=self.scrape_excel_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Export to CSV", command=self.export_to_csv).pack(side=tk.LEFT, padx=5)
        #add a button to take data from scraped file and add it to a database
        ttk.Button(buttons_frame, text="Add to Database", command=self.add_to_database).pack(side=tk.LEFT, padx=5)
        
        # Results display
        ttk.Label(self.scrape_tab, text="Scraped Data:").grid(row=4, column=0, sticky=W, padx=10, pady=5)
        
        
        # Create a Frame to contain the Text widget and scrollbars
        results_frame = ttk.Frame(self.scrape_tab)
        results_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')
        
        # Make the frame expandable
        self.scrape_tab.rowconfigure(5, weight=1)
        self.scrape_tab.columnconfigure(1, weight=1)
        
        # Create Text widget with scrollbars
        self.scrape_result_text = Text(results_frame, height=15, width=80)
        
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.scrape_result_text.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=self.scrape_result_text.xview)
        
        # Configure the Text widget to use the scrollbars
        self.scrape_result_text.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set, wrap="none")
        
        # Place widgets in the frame
        self.scrape_result_text.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Make the Text widget expandable within its frame
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

    def browse_directory(self, entry_widget):
        directory = filedialog.askdirectory()
        if directory:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, directory)
    
    def browse_file(self, entry_widget, file_types):
        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
    
    def check_unread_emails(self):
        try:
            # Clear previous results
            self.outlook_result_text.delete(1.0, tk.END)
            
            category = self.category_entry.get()
            if not category:
                messagebox.showerror("Error", "Please enter an email category.")
                return
                
            # Create OutlookProcessor instance
            processor = OutlookProcessor(
                category,
                [],  # Empty senders list for just checking
                "",  # Empty path for just checking
                ""   # Empty path for just checking
            )
            
            # Check unread emails count
            unread_count = processor.list_unread_emails()
            
            self.outlook_result_text.insert(tk.END, f"Found {unread_count} unread emails with category '{category}'.\n")
            self.status_var.set(f"Found {unread_count} unread emails")
            
        except Exception as e:
            self.outlook_result_text.insert(tk.END, f"Error: {str(e)}\n")
            self.status_var.set("Error checking emails")
    
    def process_emails(self):
        try:
            # Clear previous results
            self.outlook_result_text.delete(1.0, tk.END)
            
            category = self.category_entry.get()
            senders = [s.strip() for s in self.senders_entry.get().split(",") if s.strip()]
            attachment_path = self.attachment_path_entry.get()
            msg_path = self.msg_path_entry.get()
            
            if not category or not senders or not attachment_path or not msg_path:
                messagebox.showerror("Error", "Please fill in all required fields.")
                return
                
            self.outlook_result_text.insert(tk.END, "Starting email processing...\n")
            self.status_var.set("Processing emails...")
            self.root.update()
            
            # Create OutlookProcessor instance
            processor = OutlookProcessor(
                category,
                senders,
                attachment_path,
                msg_path
            )
            
            # Process emails in a separate thread to keep the UI responsive
            def process_thread():
                processor.download_attachments_and_save_as_msg(
                    self.save_emails_var.get(),
                    self.mark_as_read_var.get()
                )
                
                # Update UI from the main thread
                self.root.after(0, lambda: self.update_outlook_results(processor))
            
            threading.Thread(target=process_thread, daemon=True).start()
            
        except Exception as e:
            self.outlook_result_text.insert(tk.END, f"Error: {str(e)}\n")
            self.status_var.set("Error processing emails")
    
    def update_outlook_results(self, processor):
        self.outlook_result_text.insert(tk.END, "Email processing completed.\n\n")
        
        # Display results
        self.outlook_result_text.insert(tk.END, f"Emails processed: {len(processor.processed_emails)}\n")
        
        if processor.emails_with_pdf:
            self.outlook_result_text.insert(tk.END, "\nEmails with PDF attachments:\n")
            for subject in processor.emails_with_pdf:
                self.outlook_result_text.insert(tk.END, f"- {subject}\n")
                
        if processor.emails_with_nvf_new_vendor:
            self.outlook_result_text.insert(tk.END, "\nEmails with NVF or New Vendor attachments:\n")
            for subject in processor.emails_with_nvf_new_vendor:
                self.outlook_result_text.insert(tk.END, f"- {subject}\n")
        
        self.status_var.set("Email processing completed")
    
    def process_case_list(self):
        try:
            # Clear previous results
            self.case_list_result_text.delete(1.0, tk.END)
            
            excel_folder = self.excel_folder_entry.get()
            list_folder = self.list_folder_entry.get()
            
            if not excel_folder or not list_folder:
                messagebox.showerror("Error", "Please select both folders.")
                return
                
            self.case_list_result_text.insert(tk.END, "Processing case list...\n")
            self.status_var.set("Processing case list...")
            self.root.update()
            
            # Create CaseList instance
            case_list = CaseList(excel_folder, list_folder)
            
            # Process in a separate thread
            def process_thread():
                duplicate_counts, error_messages = case_list.process_excel_files()
                
                # Update UI from the main thread
                self.root.after(0, lambda: self.update_case_list_results(duplicate_counts, error_messages))
            
            threading.Thread(target=process_thread, daemon=True).start()
            
        except Exception as e:
            self.case_list_result_text.insert(tk.END, f"Error: {str(e)}\n")
            self.status_var.set("Error processing case list")
    
    def update_case_list_results(self, duplicate_counts, error_messages):
        self.case_list_result_text.insert(tk.END, "Case list processing completed.\n\n")
        
        # Show duplicates
        duplicates = sum(1 for count in duplicate_counts.values() if count > 0)
        self.case_list_result_text.insert(tk.END, f"Found {duplicates} duplicate cases.\n")
        
        if duplicates > 0:
            self.case_list_result_text.insert(tk.END, "\nDuplicate cases:\n")
            for value, count in duplicate_counts.items():
                if count > 0:
                    self.case_list_result_text.insert(tk.END, f"- {value} (Duplicated {count} times)\n")
        
        # Show errors
        if error_messages:
            self.case_list_result_text.insert(tk.END, "\nErrors encountered:\n")
            for error in error_messages:
                self.case_list_result_text.insert(tk.END, f"- {error}\n")
        
        self.status_var.set("Case list processing completed")
    
    def transpose_excel(self):
        try:
            # Clear previous results
            self.transpose_result_text.delete(1.0, tk.END)
            
            excel_file = self.excel_file_entry.get()
            sheet_name = self.sheet_name_entry.get().strip()
            
            if not excel_file:
                messagebox.showerror("Error", "Please select an Excel file.")
                return
                
            self.transpose_result_text.insert(tk.END, "Transposing Excel data...\n")
            self.status_var.set("Transposing Excel...")
            self.root.update()
            
            # Create ExcelTransposer instance
            transposer = ExcelTransposer(excel_file)
            
            # Set active sheet if specified
            if sheet_name:
                try:
                    transposer.set_active_sheet(sheet_name)
                except ValueError as e:
                    self.transpose_result_text.insert(tk.END, f"Error: {str(e)}\n")
                    self.status_var.set("Error transposing Excel")
                    return
            
            # Transpose in a separate thread
            def transpose_thread():
                transposer.transpose_cells_to_table()
                
                # Update UI from the main thread
                self.root.after(0, lambda: self.update_transpose_results(excel_file))
            
            threading.Thread(target=transpose_thread, daemon=True).start()
            
        except Exception as e:
            self.transpose_result_text.insert(tk.END, f"Error: {str(e)}\n")
            self.status_var.set("Error transposing Excel")
    
    def update_transpose_results(self, excel_file):
        self.transpose_result_text.insert(tk.END, "Excel transposition completed.\n\n")
        self.transpose_result_text.insert(tk.END, f"Transposed data saved to '{excel_file}' in a new sheet named 'Transposed'.\n")
        self.status_var.set("Excel transposition completed")
    
    def scrape_excel_files(self):
        try:
            # Clear previous results
            self.scrape_result_text.delete(1.0, tk.END)
            
            directory = self.scrape_dir_entry.get()
            range_start = self.range_start_entry.get()
            range_end = self.range_end_entry.get()
            read_headers = self.read_headers_var.get()
            
            if not directory:
                messagebox.showerror("Error", "Please select a directory with Excel files.")
                return
                
            self.scrape_result_text.insert(tk.END, f"Scraping Excel files in {directory}...\n")
            self.status_var.set("Scraping Excel files...")
            self.root.update()
            
            # Set directory for Excel scraper
            self.excel_scraper.set_directory(directory)
            
            # Scrape in a separate thread
            def scrape_thread():
                results = self.excel_scraper.scrape_excel_files(range_start, range_end, read_headers)
                
                # Update UI from the main thread
                self.root.after(0, lambda: self.update_scrape_results(results))
            
            threading.Thread(target=scrape_thread, daemon=True).start()
            
        except Exception as e:
            self.scrape_result_text.insert(tk.END, f"Error: {str(e)}\n")
            self.status_var.set("Error scraping Excel files")
    
    def update_scrape_results(self, results):
        self.scrape_result_text.delete(1.0, tk.END)
        self.scrape_result_text.insert(tk.END, f"Scraped {len(results)} Excel files.\n\n")
        
        # Display headers
        headers = self.excel_scraper.get_headers()
        if headers:
            self.scrape_result_text.insert(tk.END, "Headers found: " + ", ".join(headers) + "\n\n")
        
        # Display results in a tabular format
        if results:
            # Create header line
            header_line = f"{'Filename':<30} | "
            if headers:
                for header in headers:
                    header_line += f"{str(header):<15} | "
            else:
                # Use keys from the first result
                for key in results[0]["values"].keys():
                    header_line += f"{str(key):<15} | "
            
            self.scrape_result_text.insert(tk.END, header_line + "\n")
            self.scrape_result_text.insert(tk.END, "-" * len(header_line) + "\n")
            
            # Add each result row
            for result in results:
                line = f"{result['filename']:<30} | "
                for key, value in result["values"].items():
                    if value is None:
                        value = ""
                    line += f"{str(value):<15} | "
                self.scrape_result_text.insert(tk.END, line + "\n")
        
        self.status_var.set(f"Scraped {len(results)} Excel files")
    
    def export_to_csv(self):
        try:
            if not self.excel_scraper.get_results():
                messagebox.showerror("Error", "No data to export. Please scrape Excel files first.")
                return
                
            output_file = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            
            if not output_file:
                return  # User cancelled
                
            success = self.excel_scraper.save_results_to_csv(output_file)
            
            if success:
                messagebox.showinfo("Success", f"Data exported to {output_file}")
                self.status_var.set("Data exported to CSV")
            else:
                messagebox.showerror("Error", "Failed to export data")
                
        except Exception as e:
            messagebox.showerror("Error", f"Export error: {str(e)}")
            self.status_var.set("Error exporting data")
    
    def add_to_database(self):
        try:
            if not self.excel_scraper.get_results():
                messagebox.showerror("Error", "No data to add. Please scrape Excel files first.")
                return
                
            db_file = filedialog.asksaveasfilename(
                defaultextension=".db",
                filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")]
            )
            
            if not db_file:
                return  # User cancelled
                
            db_handler = DatabaseHandler(db_file)
            success = db_handler.add_data(self.excel_scraper.get_results())
            
            if success:
                messagebox.showinfo("Success", f"Data added to database {db_file}")
                self.status_var.set("Data added to database")
            else:
                messagebox.showerror("Error", "Failed to add data to database")
                
        except Exception as e:
            messagebox.showerror("Error", f"Database error: {str(e)}")
            self.status_var.set("Error adding data to database")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
        
