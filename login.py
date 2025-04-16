import re

from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QMessageBox, QMainWindow
import sys
import sqlite3
import bcrypt


app = QApplication(sys.argv)


def is_valid_email(email):
    return '@' in email

def is_valid_password(password):
    return (len(password) >= 8 and
    re.search(r'[A-Z]', password) and
    re.search(r'\d', password) )



def zarejestruj(window_2):
    name = name_input.text()
    email = mail_input.text()
    login = login_input.text()
    password = password_input.text()

    if not name or not email or not login or not password:
        QMessageBox.warning(window_2, "Nie podano wszystkich danych potrzebnych do logowania")
        return False



    if  not is_valid_email(email):
        QMessageBox.warning(window_2, "Błąd", "podano niepoprawny adres email")
        return False

    if not is_valid_password(password):
        QMessageBox.warning(window_2, "Błąd", "Hasło musi zawierać conajmniej 8 znaków oraz jedną dużą litere i jedną cyfre")
        return False

    password_crypted = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

    try:
        conn = sqlite3.connect('project.db')
        c = conn.cursor()

        c.execute("""CREATE TABLE IF NOT EXISTS register (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT ,
        email TEXT ,
        login TEXT ,
        password TEXT )
        """)

        c.execute("INSERT INTO register (name, email, login, password) VALUES (?,?,?,?)",(name,email,login,password_crypted))
        conn.commit()
        conn.close()
        print("Konto założone!")
        QMessageBox.information(window_2, "Sukces", "konto zostało założone!")
        return True
    except sqlite3.Error as error:
        QMessageBox.warning(window_2,"Błąd!" ,"Użytkownik już istnieje")
        return False


def register_success():
    if zarejestruj(window_2):
        window_2.hide()
        window.show()


def zaloguj():
    login = login_logowanie.text()
    password = haslo_logowanie.text()

    if not login or not password:
        QMessageBox.warning(window, "Błąd", "Musisz podać login i hasło debilu !")
        return False

    try:
        conn = sqlite3.connect('project.db')
        c = conn.cursor()
        c.execute("SELECT password FROM register WHERE login = ?",(login,))
        result = c.fetchone()
        conn.close()

        if result:
            password_crypted = result[0]
            if bcrypt.checkpw(password.encode('utf-8'), password_crypted):
                QMessageBox.information(window, "Sukces", "Zalogowano pomyślnie!")
                window.hide()
                window_3.show()
                return True
            else:
                QMessageBox.warning(window, "Błąd", "Nieprawidłowe hasłow")
                return False
        else:
            QMessageBox.warning(window, "Błąd", "Nie znaleziono użytkownika")
            return False
    except Exception as e:
        QMessageBox.critical(window, "Błąd" , f"Wystąpił błąd: {e}")



# Ekran Logowania
window = QWidget()
window.setWindowTitle("Ekran Logowania")
window.setGeometry(100, 100, 800, 600)

# Ekran Rejestracji
window_2 = QWidget()
window_2.setWindowTitle("Ekran Rejestracji")
window_2.setGeometry(100, 100, 800, 600)


# Ekran Główny aplikacji ( testowy)
window_3 = QMainWindow()
window_3.setWindowTitle("Ekran Główny")
window_3.setGeometry(100, 100, 800, 600)

# Etykiety Rejestracja
label = QLabel("Name",window_2)
label.move(370,50)
font = QFont()
font.setPointSize(17)
label.setFont(font)

label = QLabel("Email",window_2)
label.move(370,150)
font = QFont()
font.setPointSize(17)
label.setFont(font)

label = QLabel("Login",window_2)
label.move(370,250)
font = QFont()
font.setPointSize(17)
label.setFont(font)

label = QLabel("Password",window_2)
label.move(350,350)
font = QFont()
font.setPointSize(17)
label.setFont(font)

#Etykiety Logowanie
label = QLabel("Login",window)
label.move(370,100)
font = QFont()
font.setPointSize(17)
label.setFont(font)

label = QLabel("Password",window)
label.move(350,250)
font = QFont()
font.setPointSize(17)
label.setFont(font)

# Pole tekstowe Logowanie
#Login
login_logowanie = QLineEdit(window)
login_logowanie.move(300, 150)
login_logowanie.resize(200, 30)
# Hasło
haslo_logowanie = QLineEdit(window)
haslo_logowanie.move(300, 300)
haslo_logowanie.resize(200, 30)
haslo_logowanie.setEchoMode(QLineEdit.EchoMode.Password)

# Pola tekstowe Rejestracja
# Imię
name_input = QLineEdit(window_2)
name_input.move(300, 100)
name_input.resize(200, 30)
# Email
mail_input = QLineEdit(window_2)
mail_input.move(300, 200)
mail_input.resize(200, 30)
#Login
login_input = QLineEdit(window_2)
login_input.move(300, 300)
login_input.resize(200, 30)
#Hasło
password_input = QLineEdit(window_2)
password_input.setEchoMode(QLineEdit.EchoMode.Password)
password_input.move(300, 400)
password_input.resize(200, 30)

# Przyciski Rejestracja

button = QPushButton("Add user", window_2)
button.move(250, 500)
button.resize(100, 40)
button.clicked.connect(register_success)



button = QPushButton("Back", window_2)
button.move(450, 500)
button.resize(100, 40)
button.clicked.connect(window.show)
button.clicked.connect(window_2.hide)

# Przycisk Logowanie
button = QPushButton("Login", window)
button.move(250, 500)
button.resize(100, 40)
button.clicked.connect(zaloguj)

button = QPushButton("Exit", window)
button.move(450, 500)
button.resize(100, 40)
button.clicked.connect(window.close)

button_3 = QPushButton("Add new user", window)
button_3.move(350, 400)
button_3.resize(100, 40)
button_3.clicked.connect(window_2.show)
button_3.clicked.connect(window.hide)

window.show()
sys.exit(app.exec())



