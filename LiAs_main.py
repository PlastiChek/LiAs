import csv
import sys
import sqlite3
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QWidget, QTableWidgetItem, QInputDialog, QFileDialog
import openpyxl


class LiAS(QWidget):
    def __init__(self):
        super().__init__()
        # Подключение к дизайну
        uic.loadUi('librarian_helper.ui', self)

        # Подключение к базе данных
        self.con = sqlite3.connect("library.sqlite")

        # Загрузка данных
        self._load_data()

        # Соединение кнопок с функциями
        self.filter_btn.clicked.connect(self.filter_book)
        self.append_book_btn.clicked.connect(self.append_book)
        self.append_books_btn.clicked.connect(self.append_books)
        self.delete_book_btn.clicked.connect(self.delete_book)
        self.take_book_btn.clicked.connect(self.take_book)
        self.return_book_btn.clicked.connect(self.return_book)
        self.search_name_btn.clicked.connect(self.find_name)
        self.search_author_btn.clicked.connect(self.find_author)
        self.edit_btn.clicked.connect(self.edit_book)
        self.create_csv_btn.clicked.connect(self.create_csv)

    def filter_book(self):  # функция фильтра книг
        presence_condition = "presence != 'Нет'" if self.radioButton_1.isChecked() \
            else "presence = 'Нет'" if self.radioButton_2.isChecked() else None
        cur = self.con.cursor()
        if presence_condition:
            cur.execute(f"SELECT * FROM books_db WHERE {presence_condition}")
            rows = cur.fetchall()
        else:
            self._load_data()
            return
        self._update_table(rows)

    def find_name(self):  # функция поиска книг по названию
        book_name = self.search_lineedit.text()
        cur = self.con.cursor()
        cur.execute("SELECT * FROM books_db WHERE name_book = ?", (book_name,))
        rows = cur.fetchall()
        self._update_table(rows)

    def find_author(self):  # функция поиска книг по автору
        author_name = self.search_lineedit.text()
        cur = self.con.cursor()
        cur.execute("SELECT * FROM books_db WHERE author = ?", (author_name,))
        rows = cur.fetchall()
        self._update_table(rows)

    def append_book(self):  # функция добавления одной книги
        book_name, ok_pressed0 = QInputDialog.getText(self, "Добавить книгу", "Название книги")
        if ok_pressed0:
            author_name, ok_pressed1 = QInputDialog.getText(self, "Добавить книгу", "Автор книги")
            if ok_pressed1:
                presence_new = 'Нет'
                cur = self.con.cursor()
                cur.execute("INSERT INTO books_db (name_book, author, presence) VALUES (?, ?, ?)",
                            (book_name, author_name, presence_new))
                self.con.commit()
                self._load_data()

    def append_books(self):  # функция добавления книг через файл excel
        fname = QFileDialog.getOpenFileName(self, 'Выбрать excel-файл с книгами', '', 'excel-таблица (*.xlsx)')[0]
        cur = self.con.cursor()
        workbook = openpyxl.load_workbook(fname)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2):
            data = tuple(cell.value for cell in row)
            cur.execute("INSERT INTO books_db (name_book, author, presence) VALUES (?, ?, ?)",
                        (data[0], data[1], 'Нет'))
        self.con.commit()
        self._load_data()

    def delete_book(self):
        try:
            book_id, ok_pressed = QInputDialog.getInt(self, "Удалить книгу", "Введите ID книги")
            if ok_pressed:
                cur = self.con.cursor()
                cur.execute("DELETE FROM books_db WHERE id = ?", (book_id,))
                self.con.commit()
                self._load_data()
        except sqlite3.Error as e:
            print("Ошибка при удалении книги:", e)
            cur.execute("SELECT * FROM books_db WHERE id = ?", (book_id,))
            row = cur.fetchone()
            if not row:
                print("Книга с указанным ID не найдена.")
        except ValueError:
            print("Введите корректный ID книги (целое число).")

    def take_book(self):  # функция взятия книги на имя
        book_id, ok_pressed_20 = QInputDialog.getInt(self, "Взять книгу", "Введите ID книги")
        if ok_pressed_20:
            person_name, ok_pressed_21 = QInputDialog.getText(self, "Взять книгу", "Введите свою фамилию и имя")
            cur = self.con.cursor()
            cur.execute("UPDATE books_db SET presence = ? WHERE id = ?", (person_name, book_id))
            self.con.commit()
            self._load_data()

    def return_book(self):  # функция возвращения книги
        book_id, ok_pressed = QInputDialog.getInt(self, "Вернуть книгу", "Введите ID книги")
        if ok_pressed:
            cur = self.con.cursor()
            cur.execute("UPDATE books_db SET presence = ? WHERE id = ?", ('Нет', book_id))
            self.con.commit()
            self._load_data()

    def edit_book(self):  # функция редактирования данных книги
        book_id, ok_pressed_0 = QInputDialog.getInt(self, "Изменить информацию о книге", "Введите ID книги")
        if ok_pressed_0:
            book_name, ok_pressed_1 = QInputDialog.getText(self, "Изменить информацию о книге",
                                                           "Введите название книги")
            if ok_pressed_1:
                book_author, ok_pressed_2 = QInputDialog.getText(self, "Изменить информацию о книге",
                                                                 "Введите автора книги")
                if ok_pressed_2:
                    cur = self.con.cursor()
                    cur.execute("""UPDATE books_db
                        SET name_book = ?, author = ?
                        WHERE id = ?""", (book_name, book_author, book_id))
                    self.con.commit()
                    self._load_data()

    def create_csv(self):  # функция создания csv файла
        file_name, ok_pressed = QInputDialog.getText(self, "Создать csv-файл базы данных", "Введите название файла")

        if ok_pressed:
            cursor = self.con.cursor()
            cursor.execute("SELECT * FROM books_db")
            rows = cursor.fetchall()
            column_names = [description[0] for description in cursor.description]
            with open(file_name, mode='w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';', quotechar='"',
                                    quoting=csv.QUOTE_MINIMAL)
                writer.writerow(column_names)
                writer.writerows(rows)

    def _load_data(self):  # функция загрузки всех данных
        cur = self.con.cursor()
        cur.execute("SELECT * FROM books_db")
        rows = cur.fetchall()
        self._update_table(rows)

    def _update_table(self, rows):  # функция загрузки некоторых данных
        if rows:
            self.tableWidget.setRowCount(len(rows))
            self.tableWidget.setColumnCount(len(rows[0]))
            self.tableWidget.setHorizontalHeaderLabels(['ID', 'Название книги', 'Автор', 'Взята'])

            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    self.tableWidget.setItem(row_index, column_index, QTableWidgetItem(str(value)))
        else:
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(4)
            self.tableWidget.setHorizontalHeaderLabels(['ID', 'Название книги', 'Автор', 'Взята'])

    def closeEvent(self, event):  # функция закрытия бд
        self.con.close()
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = LiAS()
    ex.show()
    sys.exit(app.exec())
