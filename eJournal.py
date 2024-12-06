import sys
import sqlite3
import os
from openpyxl import Workbook
from docx import Document
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QStackedWidget, QMessageBox, QTableWidgetItem


def path_to_res(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


db_path = os.path.join(
    os.environ["USERPROFILE"], "Документы\eJournal\scheduleDB.sqlite")  # сделать обработчик ошибок и прописать вариант с Documents


class ElZhur(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi(path_to_res("elzhur.ui"), self)
        self.user = []
        self.aut.clicked.connect(self.Authorization)
        self.acc_stud.clicked.connect(self.studAc)
        self.back_stud.clicked.connect(self.Back)
        self.back.clicked.connect(self.Back_Admin)
        self.save_changes_stud.clicked.connect(self.Save_email)
        self.save_changes.clicked.connect(self.Save_admin)
        self.acc_adm.clicked.connect(self.adminAc)
        self.show_schedule.clicked.connect(self.Show_Schedule)
        self.to_export.clicked.connect(self.To_Export)
        self.to_export_2.clicked.connect(self.To_Export)
        self.back_to_shedule.clicked.connect(self.To_Schedule)
        self.create_report.clicked.connect(self.Create_Report)
        self.export_report.clicked.connect(self.Export_report)
        self.export_schedule.clicked.connect(self.Export_schedule)
        self.out.clicked.connect(self.Logout)
        self.out_admin.clicked.connect(self.Logout)

    def Logout(self):
        self.stackedWidget.setCurrentIndex(0)
        self.report.setPlainText('')
        self.user = []

    def Export_schedule(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        schedule_data = []

        if self.user[7] == 3:
            for row in range(6):
                row_data = []
                for column in range(6):
                    item = self.tableWidget_2.item(row, column)
                    item = item.text().replace('\n', ' ')
                    row_data.append(item if item else "")
                schedule_data.append(row_data)
        else:
            for row in range(6):
                row_data = []
                for column in range(6):
                    item = self.tableWidget.item(row, column)
                    item = item.text().replace('\n', ' ')
                    row_data.append(item if item else "")
                schedule_data.append(row_data)

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Расписание"

        headers = ['Понедельник', 'Вторник', 'Среда',
                   'Четверг', 'Пятница', 'Суббота']
        sheet.append(headers)

        for row in schedule_data:
            sheet.append(row)
        workbook.save(os.path.join(
            os.environ["USERPROFILE"], "Документы\schedule_export.xlsx"))  # сделать обработчик ошибок и прописать вариант с Documents

        QMessageBox.information(
            self, "Экспорт", "Расписание успешно экспортировано в schedule_export.xlsx")
        conn.close()

    def Export_report(self):
        doc = Document()
        doc.add_heading('Отчет', level=1)
        doc.add_paragraph(self.report.toPlainText())
        doc.save(os.path.join(
            os.environ["USERPROFILE"], "Документы/report.docx"))
        QMessageBox.information(
            self, "Экспорт", "Отчет успешно экспортирован в report.docx")

    def Create_Report(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        if self.user[7] == 3:
            cursor.execute("SELECT id FROM Groups WHERE group_name=?",
                           (self.groups.currentText(),))
            group = cursor.fetchone()[0]
        else:
            group = self.user[7]
        subj = self.subjects.currentText()
        if subj == "Все предметы":
            cursor.execute("""
                SELECT COUNT(*)
                FROM Schedule
                WHERE group_id=?
            """, (group,))
            count = cursor.fetchone()[0]
        else:
            cursor.execute("""
                SELECT COUNT(*)
                FROM Schedule
                WHERE group_id=? AND subject=?
            """, (group, subj))
            count = cursor.fetchone()[0]
        period = self.period.currentText()
        self.report.setPlainText(self.report.toPlainText(
        ) + f"По предмету {subj} в {period} отведено (в часах): {count * 1.5 * {'Год': 52, 'Месяц': 4,'Неделя':1}[period]}.\n")
        conn.close()

    def To_Schedule(self):
        if self.user[7] == 3:
            self.stackedWidget.setCurrentIndex(4)
        else:
            self.stackedWidget.setCurrentIndex(3)

    def To_Export(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        if self.user[7] == 3:
            cursor.execute("SELECT id FROM Groups WHERE group_name=?",
                           (self.groups.currentText(),))
            group = cursor.fetchone()
            cursor.execute(
                "SELECT subject FROM Schedule WHERE group_id=? GROUP BY subject", group)
            subjects = cursor.fetchall()
        else:
            group = self.user[7]
            cursor.execute(
                "SELECT subject FROM Schedule WHERE group_id=? GROUP BY subject", (group,))
            subjects = cursor.fetchall()
        for i in subjects:
            self.subjects.addItem(*i)
        self.stackedWidget.setCurrentIndex(5)
        conn.close()

    def Show_Schedule(self):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM Groups WHERE group_name=?",
                       (self.groups.currentText(),))
        id = cursor.fetchone()
        self.populate_schedule(*id)

    def Save_email(self):
        email = self.email.text()
        student_id = self.user[0]
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        try:
            cursor.execute(
                "UPDATE Student SET email=? WHERE id=?", (email, student_id))
            conn.commit()
            QMessageBox.information(
                self, "Успех", "Email успешно обновлен.")
            cursor.execute("SELECT * FROM Student WHERE id=?", (student_id,))
            self.user = cursor.fetchone()
            self.Back()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка сохранения",
                                f"Не удалось обновить данные: {str(e)}")
        finally:
            conn.close()

    def Back_Admin(self):
        self.stackedWidget.setCurrentIndex(4)

    def Back(self):
        self.stackedWidget.setCurrentIndex(3)

    def Authorization(self):
        log = self.login.text()
        pas = self.password.text()
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute(
            "SELECT * FROM Student WHERE login=? AND password=?", (log, pas))
        user = cursor.fetchone()

        if user:
            self.user = user
            group_id = user[7]
            if group_id == 3:
                cursor.execute(
                    "SELECT group_name FROM Groups WHERE group_name IS NOT 'Admin'")
                group_names = cursor.fetchall()
                for i in group_names:
                    self.groups.addItem(*i)
                self.stackedWidget.setCurrentIndex(4)
            else:
                cursor.execute(
                    "SELECT group_name FROM Groups WHERE id=?", (group_id,))
                group_name = cursor.fetchone()
                if group_name:
                    self.group_name.setText(group_name[0])

                self.populate_schedule(group_id)

                self.stackedWidget.setCurrentIndex(3)
        else:
            QMessageBox.warning(self, "Ошибка входа",
                                "Неверный логин или пароль.")

        conn.close()

    def studAc(self):
        self.email.setText(self.user[3])
        self.name.setText(self.user[4])
        self.surname.setText(self.user[5])
        self.patronymic.setText(self.user[6])
        self.stackedWidget.setCurrentIndex(1)

    def adminAc(self):
        self.email_admin.setText(self.user[3])
        self.name_admin.setText(self.user[4])
        self.surname_admin.setText(self.user[5])
        self.patronymic_admin.setText(self.user[6])
        self.stackedWidget.setCurrentIndex(2)

    def populate_schedule(self, group_id):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
        SELECT S.subject, S.weekday, C.class_number FROM Schedule S
        JOIN Classes C ON S.class_number = C.class_number
        WHERE S.group_id = ? ORDER BY S.weekday, C.class_number""", (group_id,))
        schedule = cursor.fetchall()
        cursor.execute("SELECT * FROM Classes")
        cl = cursor.fetchall()
        classes = {i[0]: str(i[1])+'-'+str(i[2]) for i in cl}
        schedule_dict = {day: [''] * 6 for day in ['Monday',
                                                   'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']}
        for subject, weekday, class_number in schedule:
            day_index = ['Monday', 'Tuesday', 'Wednesday',
                         'Thursday', 'Friday', 'Saturday'].index(weekday)
            schedule_dict[weekday][class_number - 1] = subject
        if self.user[7] == 3:
            self.tableWidget_2.clear()
            self.tableWidget_2.setColumnCount(6)
            self.tableWidget_2.setRowCount(6)
            self.tableWidget_2.setHorizontalHeaderLabels(
                ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'])
            for day_index, day in enumerate(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']):
                for class_index in range(6):
                    subject = classes[class_index+1] + \
                        '\n' + schedule_dict[day][class_index]
                    if subject:
                        self.tableWidget_2.setItem(
                            class_index, day_index, QTableWidgetItem(subject))
            conn.close()
            self.tableWidget_2.resizeColumnsToContents()
            for row in range(7):
                self.tableWidget_2.setRowHeight(row, 50)
        else:
            self.tableWidget.clear()
            self.tableWidget.setColumnCount(6)
            self.tableWidget.setRowCount(6)
            self.tableWidget.setHorizontalHeaderLabels(
                ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'])
            for day_index, day in enumerate(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']):
                for class_index in range(6):
                    subject = classes[class_index+1] + \
                        '\n' + schedule_dict[day][class_index]
                    if subject:
                        self.tableWidget.setItem(
                            class_index, day_index, QTableWidgetItem(subject))
                conn.close()
            self.tableWidget.resizeColumnsToContents()
            for row in range(7):
                self.tableWidget.setRowHeight(row, 50)

    def Save_admin(self):
        name = self.name_admin.text()
        surname = self.surname_admin.text()
        patronymic = self.patronymic_admin.text()
        email = self.email_admin.text()
        admin_id = self.user[0]
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        try:
            cursor.execute("UPDATE Student SET name=?, surname=?, patronymic=?, email=? WHERE id=?",
                           (name, surname, patronymic, email, admin_id))
            conn.commit()
            QMessageBox.information(
                self, "Успех", "Данные успешно обновлены.")
            cursor.execute("SELECT * FROM Student WHERE id=?", (admin_id,))
            self.user = cursor.fetchone()
            self.Back_Admin()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка сохранения",
                                f"Не удалось обновить данные: {str(e)}")
        finally:
            conn.close()


app = QApplication(sys.argv)
window = ElZhur()
window.setWindowTitle("Электронный журнал")
try:
    window.show()
except:
    pass  # сделать обработчик

app.exec()
