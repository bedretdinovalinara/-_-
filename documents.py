import sys
import sqlite3
import hashlib
import openpyxl
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QMessageBox, QStackedWidget, QComboBox, QListWidget, QDialog
)

class Database:
    def __init__(self):
        self.conn = sqlite3.connect('documents.db')
        self.cursor = self.conn.cursor()
        self.create_tables()

    def create_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY,
                full_name TEXT NOT NULL,
                username TEXT NOT NULL UNIQUE,
                password TEXT NOT NULL,
                role TEXT NOT NULL
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                content TEXT NOT NULL,
                keywords TEXT NOT NULL,
                category TEXT NOT NULL
            )
        ''')
        self.conn.commit()

    def close(self):
        self.conn.close()

    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()

    def add_user(self, full_name, username, password, role):
        hashed_password = self.hash_password(password)
        self.cursor.execute('INSERT INTO users (full_name, username, password, role) VALUES (?, ?, ?, ?)', (full_name, username, hashed_password, role))
        self.conn.commit()

    def get_user(self, username, password):
        hashed_password = self.hash_password(password)
        self.cursor.execute('SELECT role FROM users WHERE username = ? AND password = ?', (username, hashed_password))
        return self.cursor.fetchone()

    def load_documents(self):
        self.cursor.execute('SELECT id, name, content, category FROM documents')
        return self.cursor.fetchall()

    def delete_document(self, doc_id):
        self.cursor.execute('DELETE FROM documents WHERE id = ?', (doc_id,))
        self.conn.commit()

class AuthRegisterWindow(QWidget):
    def __init__(self, db, on_login):
        super().__init__()
        self.db = db
        self.on_login = on_login
        self.setWindowTitle("Авторизация и Регистрация")
        self.setGeometry(100, 100, 300, 300)

        self.layout = QVBoxLayout()
        self.stacked_widget = QStackedWidget()

        # Создаем страницы для регистрации и авторизации
        self.registration_page = self.create_registration_page()
        self.login_page = self.create_login_page()

        self.stacked_widget.addWidget(self.registration_page)
        self.stacked_widget.addWidget(self.login_page)

        self.layout.addWidget(self.stacked_widget)

        # Кнопка для переключения между страницами
        self.toggle_button = QPushButton("Уже есть аккаунт? Войти", self)
        self.toggle_button.clicked.connect(self.toggle_registration_login)
        self.layout.addWidget(self.toggle_button)

        self.setLayout(self.layout)

    def create_registration_page(self):
        page = QWidget()
        layout = QVBoxLayout()

        self.full_name_input = QLineEdit(self)
        self.full_name_input.setPlaceholderText("ФИО")
        layout.addWidget(self.full_name_input)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("Логин")
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Пароль")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input)

        # Выпадающий список для выбора роли
        self.role_combo = QComboBox(self)
        self.role_combo.addItems(["Пользователь", "Администратор"])
        layout.addWidget(self.role_combo)

        self.register_button = QPushButton("Зарегистрироваться", self)
        self.register_button.clicked.connect(self.register)
        layout.addWidget(self.register_button)

        page.setLayout(layout)
        return page

    def create_login_page(self):
        page = QWidget()
        layout = QVBoxLayout()

        self.login_username_input = QLineEdit(self)
        self.login_username_input.setPlaceholderText("Логин для входа")
        layout.addWidget(self.login_username_input)

        self.login_password_input = QLineEdit(self)
        self.login_password_input.setPlaceholderText("Пароль для входа")
        self.login_password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.login_password_input)

        self.login_button = QPushButton("Войти", self)
        self.login_button.clicked.connect(self.login)
        layout.addWidget(self.login_button)

        page.setLayout(layout)
        return page

    def toggle_registration_login(self):
        if self.stacked_widget .currentIndex() == 0:  # Если на странице регистрации
            self.stacked_widget.setCurrentIndex(1)  # Переключаем на страницу авторизации
            self.toggle_button.setText("Нет аккаунта? Зарегистрироваться")
        else:  # Если на странице авторизации
            self.stacked_widget.setCurrentIndex(0)  # Переключаем на страницу регистрации
            self.toggle_button.setText("Уже есть аккаунт? Войти")

    def register(self):
        full_name = self.full_name_input.text()
        username = self.username_input.text()
        password = self.password_input.text()
        role = self.role_combo.currentText()  # Получаем выбранную роль

        if not full_name or not username or not password:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, заполните все поля для регистрации.")
            return

        try:
            self.db.add_user(full_name, username, password, role)
            QMessageBox.information(self, "Успех", "Пользователь успешно зарегистрирован.")
            self.clear_registration_fields()
            self.stacked_widget.setCurrentIndex(1)  # Переключаем на страницу авторизации после успешной регистрации
            self.toggle_button.setText("Нет аккаунта? Зарегистрироваться")
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "Ошибка", "Имя пользователя уже существует.")

    def clear_registration_fields(self):
        self.full_name_input.clear()
        self.username_input.clear()
        self.password_input.clear()

    def login(self):
        username = self.login_username_input.text()
        password = self.login_password_input.text()
        user = self.db.get_user(username, password)

        if user:
            self.on_login(user[0])  # Передаем роль
            self.close()
        else:
            QMessageBox.warning(self, "Ошибка", "Неверное имя пользователя или пароль.")

class DocumentManagementSystem(QWidget):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.setWindowTitle("Система учета документооборота")
        self.setGeometry(100, 100, 400, 400)

        self.layout = QVBoxLayout()
        self.auth_register_window = AuthRegisterWindow(self.db, self.on_login)
        self.layout.addWidget(self.auth_register_window)

        self.setLayout(self.layout)

        self.current_role = None

    def on_login(self, role):
        self.current_role = role
        self.init_ui()

    def init_ui(self):
        if self.current_role == "Администратор":
            self.admin_ui()
        else:
            self.user_ui()

    def admin_ui(self):
        self.clear_layout()

        # Заголовок
        self.admin_label = QLabel("Управление документами и пользователями")
        self.layout.addWidget(self.admin_label)

        # Кнопки для управления документами
        self.document_management_label = QLabel("Управление документами:")
        self.layout.addWidget(self.document_management_label)

        self.add_document_button = QPushButton("Добавить документ")
        self.add_document_button.clicked.connect(self.show_add_document)
        self.layout.addWidget(self.add_document_button)

        self.view_documents_button = QPushButton("Просмотреть документы")
        self.view_documents_button.clicked.connect(self.view_documents)
        self.layout.addWidget(self.view_documents_button)

        self.delete_document_button = QPushButton("Удалить документ")
        self.delete_document_button.clicked.connect(self.delete_document)
        self.layout.addWidget(self.delete_document_button)

        # Кнопки для управления пользователями
        self.user_management_label = QLabel("Управление пользователями:")
        self.layout.addWidget(self.user_management_label)

        self.view_users_button = QPushButton("Просмотреть пользователей")
        self.view_users_button.clicked.connect(self.view_users)
        self.layout.addWidget(self.view_users_button)

        # Кнопка для экспорта отчета
        self.export_report_button = QPushButton("Создать отчёт")
        self.export_report_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_report_button)

        self.setLayout(self.layout)

    def export_to_excel(self):
        try:
            # Подключаемся к базе данных
            conn = sqlite3.connect("documents.db")  # Убедитесь, что используете правильное имя базы данных
            cursor = conn.cursor()

            # Извлекаем данные о пользователях
            cursor.execute("SELECT id, full_name, username, role FROM users")
            users = cursor.fetchall()

            if not users:
                QMessageBox.warning(self, "Нет данных", "Пользователи отсутствуют, ошибка выгрузки.")
                return

            # Создаем новый Excel файл
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Отчет по пользователям" # Записываем заголовки
            headers = ["ID", "ФИО", "Логин", "Роль"]
            sheet.append(headers)

            # Записываем данные пользователей
            for user in users:
                sheet.append(user)

            # Сохраняем файл
            file_name = "отчет_по_пользователям.xlsx"
            workbook.save(file_name)

            QMessageBox.information(self, "Успех", f"Отчет успешно экспортирован в файл: {file_name}")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при экспорте: {str(e)}")
        finally:
            # Закрываем соединение с базой данных
            if conn:
                conn.close()

    def clear_layout(self):
        # Очищаем текущий layout
        for i in reversed(range(self.layout.count())):
            widget = self.layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

    def show_add_document(self):
        # Окно для добавления документа
        self.add_document_window = AddDocumentWindow(self.db, self.load_documents)  # Передаем метод для обновления списка
        self.add_document_window.exec()  # Используем exec() для модального окна

    def view_documents(self):
        # Просмотр документов
        documents = self.db.load_documents()
        self.clear_layout()
        self.document_list = QListWidget(self)
        for doc_id, name, content, category in documents:
            self.document_list.addItem(f"{doc_id}: {name} ({category}) - {content}")
        self.layout.addWidget(self.document_list)

        self.back_button = QPushButton("Назад")
        self.back_button.clicked.connect(self.init_ui)
        self.layout.addWidget(self.back_button)

    def delete_document(self):
        # Удаление документа
        selected_item = self.document_list.currentItem()
        if selected_item:
            doc_id = int(selected_item.text().split(":")[0])
            self.db.delete_document(doc_id)
            self.view_documents()  # Обновляем список документов
        else:
            QMessageBox.warning(self, "Ошибка", "Выберите документ для удаления.")

    def view_users(self):
        # Просмотр пользователей
        self.clear_layout()
        self.user_list = QListWidget(self)
        self.db.cursor.execute('SELECT username, role FROM users')
        users = self.db.cursor.fetchall()
        for username, role in users:
            self.user_list.addItem(f"{username} - {role}")
        self.layout.addWidget(self.user_list)

        self.back_button = QPushButton("Назад")
        self.back_button.clicked.connect(self.init_ui)
        self.layout.addWidget(self.back_button)

    def user_ui(self):
        self.clear_layout()

        # Заголовок
        self.user_label = QLabel("Добро пожаловать, пользователь!")
        self.layout.addWidget(self.user_label)

        # Поле для поиска документов
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText("Поиск по названию документа")
        self.layout.addWidget(self.search_input)

        self.search_button = QPushButton("Поиск", self)
        self.search_button.clicked.connect(self.search_documents)
        self.layout.addWidget(self.search_button)

        # Список документов
        self.document_list = QListWidget(self)
        self.layout.addWidget(self.document_list)

        # Кнопка для выхода
        self.logout_button = QPushButton("Выйти")
        self.logout_button.clicked.connect(self.logout)
        self.layout.addWidget(self.logout_button)

        # Загрузка всех документов при первом входе
        self.load_documents()

    def search_documents(self):
        search_term = self.search_input.text()
        self.document_list.clear()
        documents = self.db.load_documents()
        for doc_id, name, content, category in documents:
            if search_term.lower() in name.lower():
                self.document_list.addItem(f"{doc_id}: {name} ({category}) - {content}")

    def load_documents(self):
        self.document_list.clear()
        documents = self.db.load_documents()
        for doc_id, name, content, category in documents:
            self.document_list.addItem(f"{doc_id}: {name} ({category}) - {content}")

    def logout(self):
        self.current_role = None
        self.auth_register_window.show()  # Возвращаемся к окну авторизации и регистрации

class AddDocumentWindow(QDialog):
    def __init__(self, db, update_documents_callback):
        super().__init__()
        self.db = db
        self.update_documents_callback = update_documents_callback  # Сохраняем ссылку на метод обновления
        self.setWindowTitle("Добавить документ")
        self.setGeometry(100, 100, 400, 300)

        self.layout = QVBoxLayout()

        self.name_input = QLineEdit(self)
        self.name_input.setPlaceholderText("Название документа")
        self.layout.addWidget(self.name_input)

        self.content_input = QLineEdit(self)
        self.content_input.setPlaceholderText("Содержание документа")
        self.layout.addWidget(self.content_input)

        self.keywords_input = QLineEdit(self)
        self.keywords_input.setPlaceholderText("Ключевые слова (через запятую)")
        self.layout.addWidget(self.keywords_input)

        self.category_input = QLineEdit(self)
        self.category_input.setPlaceholderText("Категория")
        self.layout.addWidget(self.category_input)

        self.add_button = QPushButton("Добавить документ", self)
        self.add_button.clicked.connect(self.add_document)
        self.layout.addWidget(self.add_button)
        self.setLayout(self.layout)

    def add_document(self):
        name = self.name_input.text()
        content = self.content_input.text()
        keywords = self.keywords_input.text()
        category = self.category_input.text()

        if not name or not content or not keywords or not category:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, заполните все поля.")
            return

        # Сохранение документа в базе данных
        self.db.cursor.execute('INSERT INTO documents (name, content, keywords, category) VALUES (?, ?, ?, ?)',
                                (name, content, keywords, category))
        self.db.conn.commit()
        QMessageBox.information(self, "Успех", "Документ успешно добавлен.")
        self.update_documents_callback()  # Обновляем список документов в главном окне
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentManagementSystem()
    window.show()
    sys.exit(app.exec())