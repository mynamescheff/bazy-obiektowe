import re
from functools import partial

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QAction
from PyQt6.QtSql import QSqlDatabase, QSqlTableModel, QSqlQuery
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QMessageBox, QMainWindow, \
    QVBoxLayout, QTableView, QInputDialog, QDialog, QFormLayout, QComboBox
import sys
import sqlite3
import bcrypt



class MainWindow(QMainWindow):
    def __init__(self, register_window, login_window):
        super().__init__()
        self.setWindowTitle("Ekran Główny")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window

        menu = self.menuBar()

        file_menu = menu.addMenu("File")
        option1_action = QAction("Outlook Processor", self)
        file_menu.addAction(option1_action)
        option2_action = QAction("Case List", self)
        file_menu.addAction(option2_action)
        option3_action = QAction("Excel Scalper", self)
        file_menu.addAction(option3_action)
        option4_action = QAction("Database Utilities", self)
        file_menu.addAction(option4_action)
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.back_or_quit)
        file_menu.addAction(exit_action)

        option_menu = menu.addMenu("Option")

        help_menu = menu.addMenu("Help")

        about_menu = menu.addMenu("About")
        version_action = QAction("Version", self)
        version_action.triggered.connect(self.version)
        about_menu.addAction(version_action)

    def back_to_login(self):
        self.hide()
        self.login_window.show()

    def version(self):
        QMessageBox.information(self, "Program", "Wersja 1.0.0")

    def back_or_quit(self):
        quit_box = QMessageBox(self)
        quit_box.setIcon(QMessageBox.Icon.Question)
        quit_box.setWindowTitle("Zamknij aplikację")
        quit_box.setText("Czy chcesz zamknąć aplikację?\nWrócić do okna logowania")

        close_button = quit_box.addButton("Zamknij aplikację", QMessageBox.ButtonRole.YesRole)
        back_button = quit_box.addButton("Powrót do logowania", QMessageBox.ButtonRole.NoRole)
        cancel_button = quit_box.addButton("Anuluj", QMessageBox.ButtonRole.RejectRole)

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
        self.setWindowTitle("Ekran Logowania")
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
            QMessageBox.warning(self, "Błąd", "Musisz podać login i hasło!")
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
                        QMessageBox.information(self, "Sukces", "Zalogowano pomyślnie!")
                        self.hide()
                        if role == "Admin":
                            self.admin_window.show()
                        else:
                            self.main_window.show()
                        return True
                else:
                    QMessageBox.warning(self, "Błąd", "Nieprawidłowe hasło")
                    return False
            else:
                QMessageBox.warning(self, "Błąd", "Nie znaleziono użytkownika")
                return False
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Wystąpił błąd: {e}")
            return False

class ChangePasswordWindow(QDialog):
    def __init__(self, login, employee_id, parent=None):
        super().__init__(parent)
        self.login = login
        self.employee_id = employee_id
        self.setWindowTitle("Zmiana hasła")
        self.setGeometry(100, 100, 400, 300)
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout()

        self.new_password_input = QLineEdit()
        self.new_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setEchoMode(QLineEdit.EchoMode.Password)

        layout.addRow("Nowe hasło:", self.new_password_input)
        layout.addRow("Potwierdź hasło:", self.confirm_password_input)

        save_button = QPushButton("Zapisz")
        save_button.clicked.connect(self.save_new_password)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Anuluj")
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
            QMessageBox.warning(self, "Błąd", "Musisz podać nowe hasło i potwierdzenie!")
            return

        if new_password != confirm_password:
            QMessageBox.warning(self, "Błąd", "Hasła nie są zgodne!")
            return

        if not self.is_valid_password(new_password):
            QMessageBox.warning(self, "Błąd", "Hasło musi zawierać co najmniej 8 znaków, jedną dużą literę i jedną cyfrę!")
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
                    QMessageBox.warning(self, "Błąd", "Nowe hasło musi różnić się od starego hasła!")
                    conn.close()
                    return


            password_crypted = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            c.execute("UPDATE Users SET password = ?, must_change_password = 0 WHERE employee_id = ?",
                      (password_crypted, self.employee_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Sukces", "Hasło zostało zmienione. Zaloguj się ponownie.")
            self.close()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się zmienić hasła: {e}")
            conn.close()

class RegisterWindow(QWidget):
    def __init__(self, login_window):
        super().__init__()
        self.setWindowTitle("Ekran Rejestracji")
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
            QMessageBox.warning(self, "Błąd", "Nie podano wszystkich danych potrzebnych do logowania")
            return False

        if password != confirm_password:
            QMessageBox.warning(self, "Błąd", "Hasła nie są zgodne")
            return False

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Błąd", "Podano niepoprawny adres email")
            return False

        if not self.is_valid_password(password):
            QMessageBox.warning(self, "Błąd",
                                "Hasło musi zawierać co najmniej 8 znaków, jedną dużą literę i jedną cyfrę")
            return False

        password_crypted = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

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

            c.execute("SELECT 1 FROM Users WHERE login = ?", (login,))
            if c.fetchone():
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim loginie już istnieje")
                conn.close()
                return False

            c.execute("SELECT 1 FROM Employees WHERE email = ?", (email,))
            if c.fetchone():
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim adresie email już istnieje")
                conn.close()
                return False

            c.execute("INSERT INTO Employees (name, second_name, email) VALUES (?, ?, ?)", (name, second_name, email))
            employee_id = c.lastrowid

            c.execute("INSERT INTO Users (login, password, employee_id) VALUES (?, ?, ?)", (login, password_crypted, employee_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Sukces", "Konto zostało założone!")
            return True
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: Users.login" in str(e):
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim loginie już istnieje")
            elif "UNIQUE constraint failed: Employees.email" in str(e):
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim adresie email już istnieje")
            else:
                QMessageBox.warning(self, "Błąd", f"Wystąpił błąd integracji: {e}")
            return False
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Błąd", f"Wystąpił błąd bazy danych: {e}")
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
        self.setWindowTitle("Ekran Administratora")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        menu = self.menuBar()
        admin_menu = menu.addMenu("Admin")
        option2_action = QAction("Refresh Data", self)
        option2_action.triggered.connect(self.register_data)
        admin_menu.addAction(option2_action)

        self.register_data()

    def register_data(self):
        if self.centralWidget():
            self.centralWidget().deleteLater()

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)


        db = QSqlDatabase.addDatabase("QSQLITE", f"admin_db_{id(self)}")
        db.setDatabaseName("project_2.db")
        if not db.open():
            QMessageBox.warning(self, "Błąd", "Nie udało się połączyć z bazą danych")
            return


        query = QSqlQuery(db)
        query.exec("PRAGMA foreign_keys = ON")
        if query.lastError().isValid():
            QMessageBox.warning(self, "Błąd", f"Nie udało się włączyć kluczy obcych: {query.lastError().text()}")
            db.close()
            return


        self.model = QSqlTableModel(self, db)
        self.model.setTable("Employees")
        self.model.setEditStrategy(QSqlTableModel.EditStrategy.OnManualSubmit)
        if not self.model.select():
            QMessageBox.warning(self, "Błąd", f"Nie udało się załadować danych: {self.model.lastError().text()}")
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
        self.setCentralWidget(central_widget)

    def add_button_to_table(self):
        delete_column = self.model.columnCount() - 2  # Przedostatnia kolumna
        modify_column = self.model.columnCount() - 1  # Ostatnia kolumna
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
        confirm = QMessageBox.question(self, "Potwierdzenie", "Czy na pewno chcesz usunąć tego użytkownika?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            employee_id = self.model.record(row).value("id")
            if self.model.removeRow(row):
                if self.model.submitAll():
                    QMessageBox.information(self, "Sukces", "Użytkownik usunięty pomyślnie.")
                    self.register_data()  # Odśwież tabelę po usunięciu
                else:
                    error = self.model.lastError().text()
                    QMessageBox.critical(self, "Błąd", f"Nie udało się usunąć użytkownika: {error}")
            else:
                QMessageBox.critical(self, "Błąd", "Nie udało się usunąć wiersza.")

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
        self.setWindowTitle("Edytuj pracownika")
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

        layout.addRow("Imię:", self.name_input)
        layout.addRow("Nazwisko:", self.second_name_input)
        layout.addRow("Email:", self.email_input)
        layout.addRow("Rola:", self.role_combo)

        save_button = QPushButton("Zapisz zmiany")
        save_button.clicked.connect(self.save_changes)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Anuluj")
        cancel_button.clicked.connect(self.close)
        layout.addWidget(cancel_button)

        self.setLayout(layout)

    def save_changes(self):
        if not self.name_input.text().strip() or not self.second_name_input.text().strip() or not self.email_input.text().strip():
            QMessageBox.warning(self, "Błąd", "Wszystkie pola muszą być wypełnione!")
            return

        if not self.is_valid_email(self.email_input.text()):
            QMessageBox.warning(self, "Błąd", "Nieprawidłowy adres email!")
            return

        self.model.setData(self.model.index(self.row, 1), self.name_input.text())  # name
        self.model.setData(self.model.index(self.row, 2), self.second_name_input.text())  # second_name
        self.model.setData(self.model.index(self.row, 3), self.email_input.text())  # email
        self.model.setData(self.model.index(self.row, 4), self.role_combo.currentText())  # role
        if self.model.submitAll():
            QMessageBox.information(self, "Sukces", "Dane pracownika zaktualizowane.")
            self.close()
        else:
            error = self.model.lastError().text()
            QMessageBox.critical(self, "Błąd", f"Nie udało się zaktualizować danych pracownika: {error}")

    def is_valid_email(self, email):
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(pattern, email) is not None

class AddNewUserWindow(QDialog):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self.setWindowTitle("Dodaj nowego użytkownika")
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

        layout.addRow("Imię:", self.name_input)
        layout.addRow("Nazwisko:", self.second_name_input)
        layout.addRow("Email:", self.email_input)
        layout.addRow("Login:", self.login_input)
        layout.addRow("Hasło:", self.password_input)
        layout.addRow("Potwierdź hasło:", self.confirm_password_input)
        layout.addRow("Rola:", self.role_combo)

        save_button = QPushButton("Dodaj użytkownika")
        save_button.clicked.connect(self.save_new_user)
        layout.addWidget(save_button)

        cancel_button = QPushButton("Anuluj")
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
            QMessageBox.warning(self, "Błąd", "Wszystkie pola muszą być wypełnione!")
            return

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Błąd", "Nieprawidłowy adres email!")
            return

        if password != confirm_password:
            QMessageBox.warning(self, "Błąd", "Hasła nie są zgodne!")
            return

        if not self.is_valid_password(password):
            QMessageBox.warning(self, "Błąd", "Hasło musi zawierać co najmniej 8 znaków, jedną dużą literę i jedną cyfrę!")
            return

        try:
            conn = sqlite3.connect('project_2.db')
            conn.execute("PRAGMA foreign_keys = ON")
            c = conn.cursor()


            c.execute("SELECT 1 FROM Users WHERE login = ?", (login,))
            if c.fetchone():
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim loginie już istnieje!")
                conn.close()
                return


            c.execute("SELECT 1 FROM Employees WHERE email = ?", (email,))
            if c.fetchone():
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim adresie email już istnieje!")
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
            QMessageBox.information(self, "Sukces", "Użytkownik został dodany pomyślnie!")
            self.close()

        except sqlite3.IntegrityError as e:
            QMessageBox.critical(self, "Błąd", f"Błąd integralności danych: {e}")
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Błąd", f"Wystąpił błąd bazy danych: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    register_window = RegisterWindow(None)
    main_window = MainWindow(register_window=register_window, login_window=None)
    admin_window = AdminWindow(register_window=register_window, login_window=None)
    login_window = LoginWindow(register_window=register_window, main_window=main_window, admin_window=admin_window)

    register_window.login_window = login_window
    main_window.login_window = login_window
    admin_window.login_window = login_window

    login_window.show()
    sys.exit(app.exec())