import os
import sqlite3
import pandas as pd

# --- Constant for the relational database ---
PROJECT_DB_PATH = "project_data.db"
CASE_LIST_DB_PATH = r".\\rozklasowany\\excelki\\cases\\case_list.db"


def _execute_query(db_path: str, query: str, params=None, text_widget_update=None):
    """Helper function to execute queries and handle common errors."""
    if not os.path.exists(db_path):
        msg = f"Error: Database '{db_path}' not found."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return None
    try:
        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return df
    except (sqlite3.OperationalError, pd.io.sql.DatabaseError) as e:
        msg = f"Database query error in '{db_path}': {e}. Ensure schema is correct and tables exist."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return None
    except Exception as e:
        msg = f"An unexpected error occurred querying '{db_path}': {e}"
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return None


def setup_project_schema(db_path: str):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS Users (
            UserID INTEGER PRIMARY KEY AUTOINCREMENT, FirstName TEXT, LastName TEXT,
            Email TEXT UNIQUE, Password TEXT );""")
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS Universities (
            UniversityID INTEGER PRIMARY KEY AUTOINCREMENT, Name TEXT UNIQUE NOT NULL,
            Address TEXT, MainBankAccount TEXT );""")
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS BankAccounts (
            BankAccountID INTEGER PRIMARY KEY AUTOINCREMENT, AccountNumber TEXT NOT NULL,
            Currency TEXT, UniversityID INTEGER,
            FOREIGN KEY (UniversityID) REFERENCES Universities(UniversityID) );""")
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS CaseNumbers (
            CaseID INTEGER PRIMARY KEY AUTOINCREMENT, CaseNumber TEXT UNIQUE NOT NULL, Title TEXT,
            Description TEXT, Status TEXT, CreationDate TEXT, LastModificationDate TEXT,
            Amount REAL, Currency TEXT, UserID INTEGER, UniversityID INTEGER,
            FOREIGN KEY (UserID) REFERENCES Users(UserID),
            FOREIGN KEY (UniversityID) REFERENCES Universities(UniversityID) );""")
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS CaseBankAccountsLink (
            CaseID INTEGER, BankAccountID INTEGER, PRIMARY KEY (CaseID, BankAccountID),
            FOREIGN KEY (CaseID) REFERENCES CaseNumbers(CaseID),
            FOREIGN KEY (BankAccountID) REFERENCES BankAccounts(BankAccountID) );""")
        conn.commit()
        print(f"Schema created/verified in '{db_path}'")
    except sqlite3.Error as e:
        print(f"Error setting up schema in '{db_path}': {e}")
    finally:
        conn.close()

def populate_project_data_from_combined_db(project_db_path: str, combined_db_path: str, text_widget_update=None):
    """
    Populates the project_data.db with data from combined.db.
    Assumes combined.db has a table 'data' with columns:
    'university', 'name', 'surname', 'case number', 'amount', 'currency', 'filename', 'bank account'.
    """
    if not os.path.exists(combined_db_path):
        msg = f"Error: Source database '{combined_db_path}' not found."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return 0

    conn_project = sqlite3.connect(project_db_path)
    cursor_project = conn_project.cursor()

    conn_combined = sqlite3.connect(combined_db_path)
    try:
        # Check if 'data' table exists in combined.db
        cursor_combined_check = conn_combined.cursor()
        cursor_combined_check.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='data';")
        if not cursor_combined_check.fetchone():
            msg = f"Table 'data' not found in {combined_db_path}"
            if text_widget_update: text_widget_update(msg)
            else: print(msg)
            return 0
        
        df_combined = pd.read_sql_query("SELECT * FROM data", conn_combined)
    except Exception as e:
        msg = f"Error reading from {combined_db_path}: {e}"
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        conn_combined.close()
        conn_project.close()
        return 0
    finally:
        conn_combined.close()

    if df_combined.empty:
        msg = f"{combined_db_path} is empty or table 'data' has no rows. No data to populate."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        conn_project.close()
        return 0

    populated_count = 0
    for index, row in df_combined.iterrows():
        try:
            # 1. University
            uni_name = str(row.get('university', 'Unknown University')).strip()
            cursor_project.execute("SELECT UniversityID FROM Universities WHERE Name = ?", (uni_name,))
            uni_result = cursor_project.fetchone()
            if uni_result:
                university_id = uni_result[0]
            else:
                cursor_project.execute("INSERT INTO Universities (Name) VALUES (?)", (uni_name,))
                university_id = cursor_project.lastrowid

            # 2. User (simplified)
            first_name = str(row.get('name', 'N/A')).strip()
            last_name = str(row.get('surname', '')).strip() # Surname might be empty
            user_email = f"{first_name.lower()}.{last_name.lower() if last_name else 'user'}@imported.example.com" if first_name != 'N/A' else "unknown.user@imported.example.com"
            
            cursor_project.execute("SELECT UserID FROM Users WHERE FirstName = ? AND LastName = ?", (first_name, last_name))
            user_result = cursor_project.fetchone()
            if user_result:
                user_id = user_result[0]
            else:
                cursor_project.execute("INSERT INTO Users (FirstName, LastName, Email) VALUES (?, ?, ?)", 
                                       (first_name, last_name, user_email))
                user_id = cursor_project.lastrowid
            
            # 3. CaseNumbers
            case_num_str = str(row.get('case number', f"UNKNOWN_CASE_{index}")).strip()
            case_title = str(row.get('filename', 'Imported Case')).strip()
            case_amount = pd.to_numeric(row.get('amount'), errors='coerce')
            case_currency = str(row.get('currency', '')).strip()
            
            # Use INSERT OR IGNORE for CaseNumbers assuming CaseNumber should be unique
            cursor_project.execute("""
                INSERT OR IGNORE INTO CaseNumbers 
                (CaseNumber, Title, Amount, Currency, UserID, UniversityID, Status, Description) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (case_num_str, case_title, case_amount, case_currency, user_id, university_id, 
                  'Imported', f'Imported from {os.path.basename(combined_db_path)}'))
            
            # Get CaseID (either newly inserted or existing if IGNORE happened)
            cursor_project.execute("SELECT CaseID FROM CaseNumbers WHERE CaseNumber = ?", (case_num_str,))
            case_id_result = cursor_project.fetchone()
            if not case_id_result:
                # This should not happen if INSERT OR IGNORE worked and CaseNumber is unique
                if text_widget_update: text_widget_update(f"Warning: Could not retrieve CaseID for {case_num_str}")
                else: print(f"Warning: Could not retrieve CaseID for {case_num_str}")
                continue 
            case_id = case_id_result[0]

            # 4. Bank Account and Link (if bank account details are present)
            bank_acc_num = str(row.get('bank account', '')).strip()
            if bank_acc_num:
                bank_acc_currency = case_currency
                
                cursor_project.execute("SELECT BankAccountID FROM BankAccounts WHERE AccountNumber = ? AND UniversityID = ?", 
                                       (bank_acc_num, university_id))
                bank_acc_result = cursor_project.fetchone()
                if bank_acc_result:
                    bank_account_id = bank_acc_result[0]
                else:
                    cursor_project.execute("INSERT INTO BankAccounts (AccountNumber, Currency, UniversityID) VALUES (?, ?, ?)",
                                           (bank_acc_num, bank_acc_currency, university_id))
                    bank_account_id = cursor_project.lastrowid
                
                # Link Case to BankAccount
                cursor_project.execute("INSERT OR IGNORE INTO CaseBankAccountsLink (CaseID, BankAccountID) VALUES (?, ?)",
                                       (case_id, bank_account_id))
            populated_count +=1
        except sqlite3.IntegrityError as ie: # Handles cases like unique constraint violations if not covered by IGNORE
            msg = f"Skipping row {index+2} due to database integrity error (e.g. duplicate unique value): {ie}"
            if text_widget_update: text_widget_update(msg)
            else: print(msg)
        except Exception as e:
            msg = f"Error processing row {index+2} from combined.db: {e}" # row index + 2 for 1-based and header
            if text_widget_update: text_widget_update(msg)
            else: print(msg)
            
    conn_project.commit()
    conn_project.close()
    
    final_msg = f"Data population from '{os.path.basename(combined_db_path)}' complete. Processed {len(df_combined)} rows. Attempted to populate/update {populated_count} entries."
    if text_widget_update: text_widget_update(final_msg)
    else: print(final_msg)
    return populated_count

def get_unique_universities_from_bank_acc_db(bank_acc_db_path: str, text_widget_update=None) -> list:
    """Fetches unique university names from bank_acc_db.db (table 'data', column 'university')."""
    if not os.path.exists(bank_acc_db_path):
        msg = f"Error: Bank account DB '{bank_acc_db_path}' not found for fetching universities."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return []
    try:
        conn = sqlite3.connect(bank_acc_db_path)
        # Check if 'data' table and 'university' column exist
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='data';")
        if not cursor.fetchone():
            raise sqlite3.OperationalError(f"Table 'data' not found in {bank_acc_db_path}")
        
        # Check for 'university' column (more robust check would be PRAGMA table_info(data))
        try:
            df = pd.read_sql_query("SELECT DISTINCT university FROM data WHERE university IS NOT NULL AND university != '' ORDER BY university", conn)
        except pd.io.sql.DatabaseError as col_err: # Typically if column doesn't exist
             raise sqlite3.OperationalError(f"Column 'university' likely missing or issue with table 'data' in {bank_acc_db_path}: {col_err}")

        conn.close()
        return sorted([str(uni) for uni in df['university'].tolist() if pd.notna(uni) and str(uni).strip()])
    except Exception as e:
        msg = f"Error fetching universities from '{bank_acc_db_path}': {e}"
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return []

def get_unique_case_numbers_from_case_list_db(case_list_db_path: str, text_widget_update=None) -> list:
    """
    Fetches unique case numbers from case_list.db.
    Assumes table 'data' and column 'case_number'.
    """
    if not os.path.exists(case_list_db_path):
        msg = f"Error: Case list DB '{case_list_db_path}' not found for fetching case numbers."
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return []
    try:
        conn = sqlite3.connect(case_list_db_path)
        # Check if 'data' table and 'case_number' column exist
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='data';")
        if not cursor.fetchone():
            raise sqlite3.OperationalError(f"Table 'data' not found in {case_list_db_path}")
        
        try:
            df = pd.read_sql_query("SELECT DISTINCT \"case_number\" FROM data WHERE \"case_number\" IS NOT NULL AND \"case_number\" != '' ORDER BY \"case_number\"", conn)
        except pd.io.sql.DatabaseError as col_err:
             raise sqlite3.OperationalError(f"Column 'case_number' likely missing or issue with table 'data' in {case_list_db_path}: {col_err}")

        conn.close()
        return sorted([str(cn) for cn in df['case_number'].tolist() if pd.notna(cn) and str(cn).strip()])
    except Exception as e:
        msg = f"Error fetching case numbers from '{case_list_db_path}': {e}"
        if text_widget_update: text_widget_update(msg)
        else: print(msg)
        return []

def display_users_with_bank_accounts(db_path: str, text_widget_update=None):
    query = """
    SELECT U.FirstName, U.LastName, U.Email, C.CaseNumber, C.Title AS CaseTitle,
           BA.AccountNumber, BA.Currency AS AccountCurrency, UniBA.Name AS BankAccountUniversity
    FROM Users U
    JOIN CaseNumbers C ON U.UserID = C.UserID
    JOIN CaseBankAccountsLink CBAL ON C.CaseID = CBAL.CaseID
    JOIN BankAccounts BA ON CBAL.BankAccountID = BA.BankAccountID
    JOIN Universities UniBA ON BA.UniversityID = UniBA.UniversityID;"""
    return _execute_query(db_path, query, text_widget_update=text_widget_update)

def display_all_cases_for_university(db_path: str, university_name: str, text_widget_update=None):
    query = """
    SELECT C.CaseNumber, C.Title, C.Description, C.Status, C.Amount, C.Currency,
           U.FirstName AS UserFirstName, U.LastName AS UserLastName
    FROM CaseNumbers C
    JOIN Universities Uni ON C.UniversityID = Uni.UniversityID
    LEFT JOIN Users U ON C.UserID = U.UserID
    WHERE Uni.Name = ?;"""
    return _execute_query(db_path, query, params=(university_name,), text_widget_update=text_widget_update)

def display_university_for_case(db_path: str, case_number_str: str, text_widget_update=None):
    query = """
    SELECT Uni.Name AS UniversityName, Uni.Address AS UniversityAddress,
           C.CaseNumber, C.Title AS CaseTitle
    FROM Universities Uni
    JOIN CaseNumbers C ON Uni.UniversityID = C.UniversityID
    WHERE C.CaseNumber = ?;"""
    return _execute_query(db_path, query, params=(case_number_str,), text_widget_update=text_widget_update)