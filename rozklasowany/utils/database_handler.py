import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Text, BooleanVar
import os
import sqlite3

class DatabaseHandler:
   def add_to_database(self):
        """
        Parse data from Excel files and add it to a SQLite database.
        The function loads Excel files, creates tables with columns matching Excel headers,
        and adds records from Excel rows to the database.
        """
        try:
            excel_data = {}
            
            # Check if there's already data loaded
            if hasattr(self, 'excel_data') and self.excel_data:
                excel_data = self.excel_data
            else:
                # Ask if user wants to select an Excel file
                response = messagebox.askyesno("No Data Found", 
                                            "No Excel data loaded. Would you like to select Excel files?")
                if response:
                    # Allow user to select Excel files
                    file_paths = filedialog.askopenfilenames(
                        title="Select Excel Files",
                        filetypes=[("Excel files", "*.xlsx *.xls")]
                    )
                    
                    if not file_paths:
                        self.status_var.set("Operation canceled")
                        return
                        
                    # Process each selected Excel file
                    for file_path in file_paths:
                        try:
                            # Read the Excel file
                            self.status_var.set(f"Reading {os.path.basename(file_path)}...")
                            
                            # Use pandas to read Excel
                            import pandas as pd
                            
                            # Read all sheets
                            excel = pd.ExcelFile(file_path)
                            file_data = {}
                            
                            for sheet_name in excel.sheet_names:
                                df = pd.read_excel(file_path, sheet_name=sheet_name)
                                
                                # Skip empty sheets
                                if df.empty:
                                    continue
                                    
                                # Convert headers to strings
                                headers = [str(col) for col in df.columns]
                                
                                # Convert data to list of lists
                                data = df.values.tolist()
                                
                                # Store the sheet data
                                sheet_key = f"{os.path.basename(file_path)}_{sheet_name}"
                                file_data[sheet_key] = {
                                    'headers': headers,
                                    'data': data
                                }
                            
                            # Store all sheets from this file
                            if file_data:
                                excel_data.update(file_data)
                                self.status_var.set(f"Successfully read {os.path.basename(file_path)}")
                            else:
                                self.status_var.set(f"No data found in {os.path.basename(file_path)}")
                                
                        except Exception as e:
                            messagebox.showwarning("Warning", f"Error reading {os.path.basename(file_path)}: {str(e)}")
                            continue
                    
                    # Store the loaded data for future use
                    self.excel_data = excel_data
                else:
                    # User declined to select files
                    self.status_var.set("Operation canceled")
                    return
                
            # Check again if we now have data
            if not excel_data:
                messagebox.showerror("Error", "No data available to add to database.")
                return
            
            # Ask user for database file location (create new or select existing)
            db_file = filedialog.asksaveasfilename(
                title="Save Database As",
                defaultextension=".db",
                filetypes=[("SQLite Database", "*.db")]
            )
            
            if not db_file:
                self.status_var.set("Database operation canceled")
                return
                
            # Connect to the SQLite database
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # Track statistics
            tables_created = 0
            records_added = 0
            
            # Process each Excel sheet's data
            for sheet_name, data in excel_data.items():
                if not data or not data.get('data') or not data.get('headers'):
                    continue
                    
                # Create a table name from the sheet name (remove extension and special chars)
                table_name = ''.join(c if c.isalnum() else '_' for c in sheet_name)
                
                # Get headers and data
                headers = data['headers']
                rows = data['data']
                
                # Create column definitions - all text fields initially for flexibility
                column_defs = [f'"{h}" TEXT' for h in headers]
                
                # Check if table exists first
                cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
                table_exists = cursor.fetchone()
                
                if not table_exists:
                    # Create the table with columns matching Excel headers
                    create_table_sql = f'CREATE TABLE "{table_name}" (id INTEGER PRIMARY KEY AUTOINCREMENT, {", ".join(column_defs)})'
                    cursor.execute(create_table_sql)
                    tables_created += 1
                    self.status_var.set(f"Created table '{table_name}' in database")
                else:
                    # Check if we need to add any new columns that exist in the Excel but not in DB
                    cursor.execute(f"PRAGMA table_info({table_name})")
                    existing_columns = [row[1] for row in cursor.fetchall()]
                    
                    for header in headers:
                        if header not in existing_columns:
                            # Add any missing columns
                            cursor.execute(f'ALTER TABLE "{table_name}" ADD COLUMN "{header}" TEXT')
                
                # Prepare and execute INSERT statements for the data
                placeholders = ', '.join(['?'] * len(headers))
                columns = ', '.join([f'"{h}"' for h in headers])
                insert_sql = f'INSERT INTO "{table_name}" ({columns}) VALUES ({placeholders})'

                
                # Insert each row of data
                for row in rows:
                    # Ensure row has data for each header (could be different lengths)
                    row_values = []
                    for h in headers:
                        idx = headers.index(h)
                        value = row[idx] if idx < len(row) else None
                        row_values.append(value)
                    
                    cursor.execute(insert_sql, row_values)
                    records_added += 1
                
                self.status_var.set(f"Added {len(rows)} records to '{table_name}' table")
            
            # Commit changes and close connection
            conn.commit()
            conn.close()
            
            messagebox.showinfo("Success", f"Database operation complete:\n- {tables_created} tables created\n- {records_added} records added\nDatabase saved as: {os.path.basename(db_file)}")
            self.status_var.set(f"Data added to database: {os.path.basename(db_file)}")
            
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"SQLite error: {str(e)}")
            self.status_var.set(f"Database error: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error adding data to database: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
