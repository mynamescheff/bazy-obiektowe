import re
from functools import partial
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QAction
from PyQt6.QtSql import QSqlDatabase, QSqlTableModel, QSqlQuery
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QMessageBox, QMainWindow, \
    QVBoxLayout, QTableView, QInputDialog, QDialog, QFormLayout, QComboBox, QTabWidget
import sys
import sqlite3
import bcrypt
from main import ExcelProcessorApp

def init_database():
    try:
        conn = sqlite3.connect('project_2.db')
        conn.execute("PRAGMA foreign_keys = ON")
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS Employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            second_name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            role TEXT DEFAULT 'Employee')
        """)
        c.execute("""CREATE TABLE IF NOT EXISTS Users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            login TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            employee_id INTEGER UNIQUE,
            must_change_password INTEGER DEFAULT 0,
            FOREIGN KEY (employee_id) REFERENCES Employees(id) ON DELETE CASCADE)
        """)
        conn.commit()
        conn.close()
    except sqlite3.Error as e:
        print(f"Database initialization error: {e}")

class MainWindow(ExcelProcessorApp):
    def __init__(self, register_window, login_window):
        super().__init__()
        self.setWindowTitle("Main Window")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window


        menu = self.menuBar()
        file_menu = menu.addMenu("File")
        option1_action = QAction("Outlook Processor", self)
        option1_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(0))
        file_menu.addAction(option1_action)
        option2_action = QAction("Case List", self)
        option2_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(1))
        file_menu.addAction(option2_action)
        option3_action = QAction("Excel Scraper", self)
        option3_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(2))
        file_menu.addAction(option3_action)
        option4_action = QAction("Database Utilities", self)
        option4_action.triggered.connect(lambda: self.tab_widget.setCurrentIndex(3))
        file_menu.addAction(option4_action)
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.back_or_quit)
        file_menu.addAction(exit_action)


        help_menu = menu.addMenu("Help")
        about_menu = menu.addMenu("About")
        version_action = QAction("Version", self)
        version_action.triggered.connect(self.version)
        about_menu.addAction(version_action)

    def back_to_login(self):
        self.hide()
        self.login_window.show()

    def version(self):
        QMessageBox.information(self, "Program", "Version 1.0.0")

    def back_or_quit(self):
        quit_box = QMessageBox(self)
        quit_box.setIcon(QMessageBox.Icon.Question)
        quit_box.setWindowTitle("Close applicantions")
        quit_box.setText("Do you want to close the application?\nReturn to the login window")

        close_button = quit_box.addButton("Close applicantions", QMessageBox.ButtonRole.YesRole)
        back_button = quit_box.addButton("Return to the login window", QMessageBox.ButtonRole.NoRole)
        cancel_button = quit_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)

        quit_box.exec()

        if quit_box.clickedButton() == close_button:
            QApplication.quit()
        elif quit_box.clickedButton() == back_button:
            self.back_to_login()
        else:
            pass

class LoginWindow(QWidget):
    def __init__(self, register_window, main_window, admin_window):
        super().__init__()
        self.setWindowTitle("Login Window")
        self.setGeometry(100, 100, 800, 600)
        self.main_window = main_window
        self.register_window = register_window
        self.admin_window = admin_window
        self.LoginUI()

    def LoginUI(self):
        label = QLabel("Login", self)
        label.move(370, 100)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.login_logowanie = QLineEdit(self)
        self.login_logowanie.move(300, 150)
        self.login_logowanie.resize(200, 30)

        label = QLabel("Password", self)
        label.move(350, 250)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.haslo_logowanie = QLineEdit(self)
        self.haslo_logowanie.move(300, 300)
        self.haslo_logowanie.resize(200, 30)
        self.haslo_logowanie.setEchoMode(QLineEdit.EchoMode.Password)

        button = QPushButton("Login", self)
        button.move(250, 500)
        button.resize(100, 40)
        button.clicked.connect(self.zaloguj)

        button = QPushButton("Exit", self)
        button.move(450, 500)
        button.resize(100, 40)
        button.clicked.connect(self.close)

        button_3 = QPushButton("Add new user", self)
        button_3.move(350, 400)
        button_3.resize(100, 40)
        button_3.clicked.connect(self.go_to_register)

    def go_to_register(self):
        self.hide()
        self.register_window.show()

    def zaloguj(self):
        login = self.login_logowanie.text()
        password = self.haslo_logowanie.text()

        if not login or not password:
            QMessageBox.warning(self, "Error", "You must enter your login and password!")
            return False

        try:
            conn = sqlite3.connect('project_2.db')
            c = conn.cursor()
            c.execute("SELECT Users.password, Employees.role, Users.employee_id, Users.must_change_password FROM Users JOIN Employees ON Users.employee_id = Employees.id WHERE Users.login = ?", (login,))
            result = c.fetchone()
            conn.close()

            if result:
                password_crypted, role, employee_id, must_change_password = result
                if bcrypt.checkpw(password.encode('utf-8'), password_crypted.encode('utf-8')):
                    if must_change_password:
                        self.change_password_window = ChangePasswordWindow(login, employee_id, self)
                        self.change_password_window.show()
                    else:
                        QMessageBox.information(self, "Success”, ”Logged in successfully!")
                        self.hide()
                        if role == "Admin":
                            self.admin_window.show()
                        else:
                            self.main_window.show()
                        return True
                else:
                    QMessageBox.warning(self, "Error", "Incorrect password")
                    return False
            else:
                QMessageBox.warning(self, "Error", "User not found")
                return False
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error has occurred: {e}")
            return False

class ChangePasswordWindow(QDialog):
    def __init__(self, login, employee_id, parent=None):
        super().__init__(parent)
        self.login = login
        self.employee_id = employee_id
        self.setWindowTitle("Change password")
        self.setGeometry(100, 100, 400, 300)
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()

        self.new_password_input = QLineEdit()
        self.new_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setEchoMode(QLineEdit.EchoMode.Password)

        layout.addRow("New password:", self.new_password_input)
        layout.addRow("Confirm password:", self.confirm_password_input)

        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_new_password)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)
        layout.addWidget(cancel_button)

        self.setLayout(layout)

    def is_valid_password(self, password):
        return (len(password) >= 8 and
                re.search(r'[A-Z]', password) and
                re.search(r'\d', password))

    def save_new_password(self):
        new_password = self.new_password_input.text().strip()
        confirm_password = self.confirm_password_input.text().strip()

        if not new_password or not confirm_password:
            QMessageBox.warning(self, "Error", "You must enter a new password and confirmation!")
            return

        if new_password != confirm_password:
            QMessageBox.warning(self, "Error", "The passwords do not match!")
            return

        if not self.is_valid_password(new_password):
            QMessageBox.warning(self, "Error", "The password must contain at least 8 characters, one capital letter, and one number!")
            return

        try:
            conn = sqlite3.connect('project_2.db')
            conn.execute("PRAGMA foreign_keys = ON")
            c = conn.cursor()

            c.execute("SELECT password FROM Users WHERE employee_id = ?", (self.employee_id,))
            result = c.fetchone()
            if result:
                old_password_crypted = result[0]
                if bcrypt.checkpw(new_password.encode('utf-8'), old_password_crypted.encode('utf-8')):
                    QMessageBox.warning(self, "Error", "The new password must be different from the old password!")
                    conn.close()
                    return

            password_crypted = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            c.execute("UPDATE Users SET password = ?, must_change_password = 0 WHERE employee_id = ?",
                      (password_crypted, self.employee_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", "Your password has been changed. Please log in again.")
            self.close()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"Unable to change password: {e}")
            conn.close()

class RegisterWindow(QWidget):
    def __init__(self, login_window):
        super().__init__()
        self.setWindowTitle("Registration Screen")
        self.setGeometry(100, 100, 800, 700)
        self.login_window = login_window
        self.RegisterUI()

    def RegisterUI(self):
        label = QLabel("Name", self)
        label.move(370, 50)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.name_input = QLineEdit(self)
        self.name_input.move(300, 100)
        self.name_input.resize(200, 30)

        label = QLabel("Second name", self)
        label.move(330, 150)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.second_name = QLineEdit(self)
        self.second_name.move(300, 200)
        self.second_name.resize(200, 30)

        label = QLabel("Email", self)
        label.move(370, 250)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.mail_input = QLineEdit(self)
        self.mail_input.move(300, 300)
        self.mail_input.resize(200, 30)

        label = QLabel("Login", self)
        label.move(370, 350)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.login_input = QLineEdit(self)
        self.login_input.move(300, 400)
        self.login_input.resize(200, 30)

        label = QLabel("Password", self)
        label.move(350, 450)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.move(300, 500)
        self.password_input.resize(200, 30)

        label = QLabel("Confirm Password", self)
        label.move(310, 550)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.confirm_password_input = QLineEdit(self)
        self.confirm_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_password_input.move(300, 600)
        self.confirm_password_input.resize(200, 30)

        button = QPushButton("Add user", self)
        button.move(250, 650)
        button.resize(100, 40)
        button.clicked.connect(self.register_success)

        button = QPushButton("Back", self)
        button.move(450, 650)
        button.resize(100, 40)
        button.clicked.connect(self.back_to_login)

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email) is not None

    def is_valid_password(self, password):
        return (len(password) >= 8 and
                re.search(r'[A-Z]', password) and
                re.search(r'\d', password))

    def zarejestruj(self):
        name = self.name_input.text().strip()
        second_name = self.second_name.text().strip()
        email = self.mail_input.text().strip()
        login = self.login_input.text().strip()
        password = self.password_input.text().strip()
        confirm_password = self.confirm_password_input.text().strip()

        if not name or not second_name or not email or not login or not password or not confirm_password:
            QMessageBox.warning(self, "Error", "Not all the data required for login has been provided")
            return False

        if password != confirm_password:
            QMessageBox.warning(self, "Error", "Passwords do not match")
            return False

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Error", "Invalid email address provided")
            return False

        if not self.is_valid_password(password):
            QMessageBox.warning(self, "Error",
                                "The password must contain at least 8 characters, one capital letter, and one number")
            return False

        password_crypted = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

        try:
            conn = sqlite3.connect('project_2.db')
            conn.execute("PRAGMA foreign_keys = ON")
            c = conn.cursor()

            c.execute("SELECT 1 FROM Users WHERE login = ?", (login,))
            if c.fetchone():
                QMessageBox.warning(self, "Error", "A user with this login already exists")
                conn.close()
                return False

            c.execute("SELECT 1 FROM Employees WHERE email = ?", (email,))
            if c.fetchone():
                QMessageBox.warning(self, "Error", "A user with this email address already exists")
                conn.close()
                return False

            c.execute("INSERT INTO Employees (name, second_name, email) VALUES (?, ?, ?)", (name, second_name, email))
            employee_id = c.lastrowid

            c.execute("INSERT INTO Users (login, password, employee_id) VALUES (?, ?, ?)", (login, password_crypted, employee_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Success", "Your account has been created!")
            return True
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: Users.login" in str(e):
                QMessageBox.warning(self, "Error", "A user with this login already exists")
            elif "UNIQUE constraint failed: Employees.email" in str(e):
                QMessageBox.warning(self, "Error", "A user with this email address already exists")
            else:
                QMessageBox.warning(self, "Error", f"An integration error has occurred: {e}")
            return False
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Error", f"A database error has occurred: {e}")
            return False

    def register_success(self):
        if self.zarejestruj():
            self.hide()
            self.login_window.show()

    def back_to_login(self):
        self.hide()
        self.login_window.show()

class AdminWindow(MainWindow):
    def __init__(self, register_window, login_window):
        super().__init__(register_window, login_window)
        self.setWindowTitle("Administrator Window")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window

        self.admin_tab = QWidget()
        self.tab_widget.addTab(self.admin_tab, "Admin")
        self.tab_widget.setCurrentWidget(self.admin_tab)


        self.register_data()

        menu = self.menuBar()
        admin_menu = menu.addMenu("Admin")
        option2_action = QAction("Refresh Data", self)
        option2_action.triggered.connect(self.register_data)
        admin_menu.addAction(option2_action)

        self.register_data()

    def register_data(self):
        layout = QVBoxLayout(self.admin_tab)
        print("Test")

        db = QSqlDatabase.addDatabase("QSQLITE", f"admin_db_{id(self)}")
        db.setDatabaseName("project_2.db")
        if not db.open():
            QMessageBox.warning(self, "Error", "Unable to connect to the database")
            return

        query = QSqlQuery(db)
        query.exec("PRAGMA foreign_keys = ON")
        if query.lastError().isValid():
            QMessageBox.warning(self, "Error", f"Unable to activate foreign keys: {query.lastError().text()}")
            db.close()
            return

        self.model = QSqlTableModel(self, db)
        self.model.setTable("Employees")
        self.model.setEditStrategy(QSqlTableModel.EditStrategy.OnManualSubmit)
        if not self.model.select():
            QMessageBox.warning(self, "Error", f"Failed to load data: {self.model.lastError().text()}")
            db.close()
            return

        column_count = self.model.columnCount()
        self.model.insertColumn(column_count)
        self.model.setHeaderData(column_count, Qt.Orientation.Horizontal, "Delete")
        self.model.insertColumn(column_count + 1)
        self.model.setHeaderData(column_count + 1, Qt.Orientation.Horizontal, "Modify")

        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.setColumnHidden(0, True)

        self.add_button_to_table()

        layout.addWidget(self.table_view)
        self.add_button = QPushButton("Add new user")
        self.add_button.clicked.connect(self.add_new_user)
        layout.addWidget(self.add_button)

    def add_button_to_table(self):
        delete_column = self.model.columnCount() - 2
        modify_column = self.model.columnCount() - 1
        for row in range(self.model.rowCount()):
            delete_button = QPushButton("Delete")
            delete_button.setStyleSheet("background-color: red; color: white;")
            delete_button.clicked.connect(partial(self.delete_record, row))

            modify_button = QPushButton("Modify")
            modify_button.setStyleSheet("background-color: blue; color: white;")
            modify_button.clicked.connect(partial(self.edit_employee, row))

            self.table_view.setIndexWidget(self.model.index(row, delete_column), delete_button)
            self.table_view.setIndexWidget(self.model.index(row, modify_column), modify_button)

    def delete_record(self, row):
        confirm = QMessageBox.question(self, "Confirmation", "Are you sure you want to delete this user?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            employee_id = self.model.record(row).value("id")
            if self.model.removeRow(row):
                if self.model.submitAll():
                    QMessageBox.information(self, "Success", "User successfully deleted.")
                    self.register_data()
                else:
                    error = self.model.lastError().text()
                    QMessageBox.critical(self, "Error", f"Unable to delete user: {error}")
            else:
                QMessageBox.critical(self, "Error", "The row could not be deleted.")

    def edit_employee(self, row):
        record = self.model.record(row)
        employee_data = {
            "id": record.value("id"),
            "name": record.value("name"),
            "second_name": record.value("second_name"),
            "email": record.value("email"),
            "role": record.value("role")
        }
        self.edit_employee_window = EditEmployeeWindow(employee_data, self.model, row, self)
        self.edit_employee_window.show()

    def add_new_user(self):
        self.add_new_user_window = AddNewUserWindow(self.model, self)
        self.add_new_user_window.show()

class EditEmployeeWindow(QDialog):
    def __init__(self, employee_data, model, row, parent=None):
        super().__init__(parent)
        self.employee_data = employee_data
        self.model = model
        self.row = row
        self.setWindowTitle("Edit employee")
        self.setGeometry(100, 100, 400, 300)
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()

        self.name_input = QLineEdit(self.employee_data['name'])
        self.second_name_input = QLineEdit(self.employee_data['second_name'])
        self.email_input = QLineEdit(self.employee_data['email'])
        self.role_combo = QComboBox()
        self.role_combo.addItems(["Employee", "Admin"])
        self.role_combo.setCurrentText(self.employee_data['role'])

        layout.addRow("Name:", self.name_input)
        layout.addRow("Second name:", self.second_name_input)
        layout.addRow("Email:", self.email_input)
        layout.addRow("Role:", self.role_combo)

        save_button = QPushButton("Save changes")
        save_button.clicked.connect(self.save_changes)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)
        layout.addWidget(cancel_button)

        self.setLayout(layout)

    def save_changes(self):
        if not self.name_input.text().strip() or not self.second_name_input.text().strip() or not self.email_input.text().strip():
            QMessageBox.warning(self, "Error", "All fields must be filled in!")
            return

        if not self.is_valid_email(self.email_input.text()):
            QMessageBox.warning(self, "Error", "Invalid email address!")
            return

        self.model.setData(self.model.index(self.row, 1), self.name_input.text())
        self.model.setData(self.model.index(self.row, 2), self.second_name_input.text())
        self.model.setData(self.model.index(self.row, 3), self.email_input.text())
        self.model.setData(self.model.index(self.row, 4), self.role_combo.currentText())
        if self.model.submitAll():
            QMessageBox.information(self, "Success", "Employee data updated.")
            self.close()
        else:
            error = self.model.lastError().text()
            QMessageBox.critical(self, "Error", f"Failed to update employee data: {error}")

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email) is not None

class AddNewUserWindow(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self.setWindowTitle("Add new user")
        self.setGeometry(100, 100, 400, 400)
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()

        self.name_input = QLineEdit()
        self.second_name_input = QLineEdit()
        self.email_input = QLineEdit()
        self.login_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.role_combo = QComboBox()
        self.role_combo.addItems(["Employee", "Admin"])

        layout.addRow("Name:", self.name_input)
        layout.addRow("Second name:", self.second_name_input)
        layout.addRow("Email:", self.email_input)
        layout.addRow("Login:", self.login_input)
        layout.addRow("Password:", self.password_input)
        layout.addRow("Confirm password:", self.confirm_password_input)
        layout.addRow("Role:", self.role_combo)

        save_button = QPushButton("Add user")
        save_button.clicked.connect(self.save_new_user)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)
        layout.addWidget(cancel_button)

        self.setLayout(layout)

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email) is not None

    def is_valid_password(self, password):
        return (len(password) >= 8 and
                re.search(r'[A-Z]', password) and
                re.search(r'\d', password))

    def save_new_user(self):
        name = self.name_input.text().strip()
        second_name = self.second_name_input.text().strip()
        email = self.email_input.text().strip()
        login = self.login_input.text().strip()
        password = self.password_input.text().strip()
        confirm_password = self.confirm_password_input.text().strip()
        role = self.role_combo.currentText()

        if not all([name, second_name, email, login, password, confirm_password]):
            QMessageBox.warning(self, "Error", "All fields must be filled in!")
            return

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Error", "Invalid email address!")
            return

        if password != confirm_password:
            QMessageBox.warning(self, "Error", "The passwords do not match!")
            return

        if not self.is_valid_password(password):
            QMessageBox.warning(self, "Error", "The password must contain at least 8 characters, one capital letter, and one number!")
            return

        try:
            conn = sqlite3.connect('project_2.db')
            conn.execute("PRAGMA foreign_keys = ON")
            c = conn.cursor()

            c.execute("SELECT 1 FROM Users WHERE login = ?", (login,))
            if c.fetchone():
                QMessageBox.warning(self, "Error", "A user with this login already exists!")
                conn.close()
                return

            c.execute("SELECT 1 FROM Employees WHERE email = ?", (email,))
            if c.fetchone():
                QMessageBox.warning(self, "Error", "A user with this email address already exists!")
                conn.close()
                return

            c.execute("INSERT INTO Employees (name, second_name, email, role) VALUES (?, ?, ?, ?)",
                      (name, second_name, email, role))
            employee_id = c.lastrowid

            password_crypted = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            c.execute("INSERT INTO Users (login, password, employee_id, must_change_password) VALUES (?, ?, ?, ?)",
                      (login, password_crypted, employee_id, 1))
            conn.commit()
            conn.close()

            self.model.select()
            QMessageBox.information(self, "Success", "User added successfully!")
            self.close()

        except sqlite3.IntegrityError as e:
            QMessageBox.critical(self, "Error", f"Data integrity error: {e}")
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Error", f"A database error has occurred: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    init_database()
    register_window = RegisterWindow(None)
    main_window = MainWindow(register_window=register_window, login_window=None)
    admin_window = AdminWindow(register_window=register_window, login_window=None)
    login_window = LoginWindow(register_window=register_window, main_window=main_window, admin_window=admin_window)

    register_window.login_window = login_window
    main_window.login_window = login_window
    admin_window.login_window = login_window

    login_window.show()
    sys.exit(app.exec())