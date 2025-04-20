import re
from functools import partial

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QAction
from PyQt6.QtSql import QSqlDatabase, QSqlTableModel
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QMessageBox, QMainWindow, \
    QVBoxLayout, QTableView, QInputDialog, QStyledItemDelegate, QToolButton, QDialog
import sys
import sqlite3
import bcrypt
from pyexpat import model


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
            QMessageBox.warning(self, "Błąd", "Musisz podać login i hasło debilu !")
            return False

        try:
            conn = sqlite3.connect('project.db')
            c = conn.cursor()
            c.execute("SELECT password FROM register WHERE login = ?", (login,))
            result = c.fetchone()
            conn.close()

            if result:
                password_crypted = result[0]
                if bcrypt.checkpw(password.encode('utf-8'), password_crypted):
                    QMessageBox.information(self, "Sukces", "Zalogowano pomyślnie!")
                    self.hide()

                    if login == "admin" and password == "Administrator1":
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


class RegisterWindow(QWidget):
    def __init__(self, login_window):
        super().__init__()
        self.setWindowTitle("Ekran Rejestracji")
        self.setGeometry(100, 100, 800, 600)
        self.RegisterUI()
        self.login_window = login_window



    def RegisterUI(self):

        label = QLabel("Name", self)
        label.move(370, 50)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)


        self.name_input = QLineEdit(self)
        self.name_input.move(300, 100)
        self.name_input.resize(200, 30)


        label = QLabel("Email", self)
        label.move(370, 150)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.mail_input = QLineEdit(self)
        self.mail_input.move(300, 200)
        self.mail_input.resize(200, 30)

        label = QLabel("Login", self)
        label.move(370, 250)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.login_input = QLineEdit(self)
        self.login_input.move(300, 300)
        self.login_input.resize(200, 30)

        label = QLabel("Password", self)
        label.move(350, 350)
        font = QFont()
        font.setPointSize(17)
        label.setFont(font)

        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.move(300, 400)
        self.password_input.resize(200, 30)

        button = QPushButton("Add user", self)
        button.move(250, 500)
        button.resize(100, 40)
        button.clicked.connect(self.register_success)

        button = QPushButton("Back", self)
        button.move(450, 500)
        button.resize(100, 40)
        button.clicked.connect(self.back_to_login)

    def is_valid_email(self, email):
        return '@' in email

    def is_valid_password(self, password):
        return (len(password) >= 8 and
                re.search(r'[A-Z]', password) and
                re.search(r'\d', password))

    def zarejestruj(self):
        name = self.name_input.text()
        email = self.mail_input.text()
        login = self.login_input.text()
        password = self.password_input.text()

        if not name or not email or not login or not password:
            QMessageBox.warning(self, "Błąd","Nie podano wszystkich danych potrzebnych do logowania")
            return False

        if not self.is_valid_email(email):
            QMessageBox.warning(self, "Błąd", "podano niepoprawny adres email")
            return False

        if not self.is_valid_password(password):
            QMessageBox.warning(self, "Błąd",
                                "Hasło musi zawierać conajmniej 8 znaków oraz jedną dużą litere i jedną cyfre")
            return False

        password_crypted = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

        try:
            conn = sqlite3.connect('project.db')
            c = conn.cursor()

            c.execute("""CREATE TABLE IF NOT EXISTS register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT ,
            email TEXT UNIQUE ,
            login TEXT UNIQUE ,
            password TEXT )
            """)

            c.execute("SELECT 1 FROM register WHERE login = ? ",(login,))
            if c.fetchone():
                QMessageBox.warning(self, "Błąd", "Użytkownik o takim loginie już istnieje")
                conn.close()
                return False

            c.execute("SELECT 1 FROM register WHERE email = ? ",(email,))
            if c.fetchone():
                QMessageBox.warning(self, "Bląd", "Użytkownik o takim adresu email już istnieje")
                conn.close()
                return False


            c.execute("INSERT INTO register (name, email, login, password) VALUES (?,?,?,?)",
                      (name, email, login, password_crypted))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "Sukces", "konto zostało założone!")
            return True
        except sqlite3.Error as error:
            QMessageBox.warning(self, "Błąd!", f"Wystąpił błąd: {error}: ")
            return False

    def register_success(self):
        if self.zarejestruj():
            self.hide()
            self.login_window.show()

    def back_to_login(self):
        self.hide()
        self.login_window.show()



class MainWindow(QMainWindow):
    def __init__(self, register_window, login_window):
        super().__init__()
        self.setWindowTitle("Ekran Główny")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window

        menu = self.menuBar()

        file_menu = menu.addMenu("File")
        option1_action = QAction("Test1", self)
        file_menu.addAction(option1_action)
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
        quit_box.setWindowTitle("Zamknij aplikacje")
        quit_box.setText("Czy chcesz zamknąć aplikacje ?\n Wrócić do okna logowania")

        close_button = quit_box.addButton("Zakmnij aplikacje", QMessageBox.ButtonRole.YesRole)
        back_button  = quit_box.addButton("Powrót do logowania", QMessageBox.ButtonRole.NoRole)
        cancel_button = quit_box.addButton("Anuluj", QMessageBox.ButtonRole.RejectRole)

        quit_box.exec()

        if  quit_box.clickedButton() == close_button:
            QApplication.quit()
        elif quit_box.clickedButton() == back_button:
            self.back_to_login()
        else:
            pass









class AdminWindow(MainWindow):
    def __init__(self, register_window, login_window):
        super().__init__(register_window, login_window)
        self.setWindowTitle("Ekran Administrator")
        self.setGeometry(100, 100, 800, 600)
        self.register_window = register_window
        self.login_window = login_window

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        menu = self.menuBar()
        admin_menu = menu.addMenu("Admin")
        option2_action = QAction("Test2", self)
        option2_action.triggered.connect(self.register_data)
        admin_menu.addAction(option2_action)

    def register_data(self):
        if self.centralWidget():
            self.centralWidget().deleteLater()

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)

        db = QSqlDatabase.addDatabase("QSQLITE")
        if not db.isOpen():
            db.setDatabaseName("project.db")
            if not db.open():
                QMessageBox.warning(self, "Błąd", "Nie udało się połączyć z bazą danych")
                return

        self.model = QSqlTableModel(self, db)
        self.model.setTable("register")
        self.model.select()

        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.setColumnHidden(0, True)


        if self.model.columnCount() <= 5:
            self.model.insertColumn(self.model.columnCount())
            self.model.setHeaderData(self.model.columnCount() - 1, Qt.Orientation.Horizontal, "Delete")

            self.model.insertColumn(self.model.columnCount())
            self.model.setHeaderData(self.model.columnCount() - 1, Qt.Orientation.Horizontal, "Edit")


        self.add_button()

        layout.addWidget(self.table_view)
        self.setCentralWidget(central_widget)

    def add_button(self):
        delete_column = self.model.columnCount() - 2
        edit_column = self.model.columnCount() - 1


        for row in range(self.model.rowCount()):
            delete_button = QPushButton("Delete")
            delete_button.setStyleSheet("background-color: red; color: white;")
            delete_button.clicked.connect(partial(self.delete_record, row))

            edit_button = QPushButton("Edit")
            edit_button.setStyleSheet("background-color: blue; color: white;")
            # edit_button.clicked.connect(partial(self.edit_record, row))


            self.table_view.setIndexWidget(self.model.index(row, delete_column), delete_button)
            self.table_view.setIndexWidget(self.model.index(row, edit_column), edit_button)

    def delete_record(self, row):
        confirm = QMessageBox.question(self, "Potwierdzenie", "Czy na pewno chcesz usunąć tego użytkownika?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if confirm == QMessageBox.StandardButton.Yes:
            self.model.removeRow(row)
            self.model.submitAll()
            self.register_data()

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







