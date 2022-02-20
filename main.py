import sys
import sqlite3
import os
import datetime as dt

from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QTableWidgetItem
from PyQt5.QtWidgets import QMessageBox, QInputDialog


class DateError(Exception):
    pass


class OlympError(Exception):
    pass


class GradeError(Exception):
    pass


class SubjectError(Exception):
    pass


def cor_birthday(birth):
    if len(birth) != 10:
        return False
    elif (not birth[:4].isdigit() or not birth[5:7].isdigit() or not birth[8:].isdigit()) and \
            (not birth[:2].isdigit() or not birth[3:5].isdigit() or not birth[6:].isdigit()):
        return False
    return True


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("interface.ui", self)
        self.preview = None
        self.flag1 = False
        self.flag2 = False
        self.flag3 = False
        self.flag4 = False
        self.flag5 = False
        self.Printe.clicked.connect(self.printe)
        self.Printe2.clicked.connect(self.printe2)
        self.Printe3.clicked.connect(self.printe3)
        self.Printe4.clicked.connect(self.printe4)
        self.AppendSt.clicked.connect(self.append_student)
        self.DeleteSt.clicked.connect(self.delete_student)
        self.AppendOlymp.clicked.connect(self.append_olymp)
        self.DeleteOlymp.clicked.connect(self.delete_olymp)
        self.AppendOlympR.clicked.connect(self.append_olymp_r)
        self.DeleteOlympR.clicked.connect(self.delete_olymp_r)
        self.Ta.itemChanged.connect(self.change1)
        self.Ta2.itemChanged.connect(self.change2)
        self.Ta3.itemChanged.connect(self.change3)
        self.Ta4.itemChanged.connect(self.change4)
        self.Ta5.itemChanged.connect(self.change5)
        self.Class1.currentTextChanged.connect(self.class1_change)
        self.Class2.currentTextChanged.connect(self.class2_change)
        self.Class3.currentTextChanged.connect(self.run3)
        self.Class4.currentTextChanged.connect(self.class4_change)
        self.Class5.currentTextChanged.connect(self.run5)
        self.Olymps.currentTextChanged.connect(self.run4)
        self.Subjects1.currentTextChanged.connect(self.run1)
        self.olymps_change()
        self.chetvs = {1: [dt.date(2021, 9, 1), 65]}
        self.class1_change()
        self.class2_change()
        self.class4_change()
        self.run1()
        self.run2()
        self.run3()
        self.run4()
        self.run5()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Right:
            name = self.TAB.currentWidget().objectName()
            if name == "tab5":
                ind = 0
            else:
                ind = int(name[-1]) + 1 - 1
            print(ind)
            a = [self.tab1, self.tab2, self.tab3, self.tab4, self.tab5]
            self.TAB.setCurrentWidget(a[ind])
        elif event.key() == Qt.Key_Left:
            name = self.TAB.currentWidget().objectName()
            if name == "tab1":
                ind = 4
            else:
                ind = int(name[-1]) - 1 - 1
            print(ind)
            a = [self.tab1, self.tab2, self.tab3, self.tab4, self.tab5]
            self.TAB.setCurrentWidget(a[ind])

    def change1(self):
        # Эта функция отвечает за изменение выделеных оценок в таблице.
        # Здесь нужен флаг,
        # чтобы функция не вызывалась лишний раз при создании таблицы.
        if not self.flag1:
            return
        con = sqlite3.connect("Study.db")
        x, y = [(i.column(), i.row()) for i in self.Ta.selectedItems()][0]
        try:
            date = self.Ta.horizontalHeaderItem(x).text()
            name = self.Ta.verticalHeaderItem(y).text()
            subject = self.Subjects1.currentText()
            cur = con.cursor()
            ind = cur.execute(f"""select id 
from Students where name = '{name}'""").fetchone()[0]
            ind2 = cur.execute(f"""select id 
from Subjects where name = '{subject}'""").fetchone()[0]
            if self.Ta.item(y, x) and self.Ta.item(y, x).text().isdigit():
                grade = int(self.Ta.item(y, x).text())
            else:
                cur.execute(f"""delete from Grades
                where student = (select id from Students where name = '{name}') 
                and date = '{date}' and subject = {ind2}""")
                con.commit()
                # Здесь нужен флаг чтобы не возникла рекурсия
                self.flag1 = False
                self.Ta.setItem(y, x, QTableWidgetItem(""))
                self.flag1 = True
                return
            if grade > 5 or grade < 0:
                raise GradeError
            # Сначала удаляется прошлое значение оценки
            # если оно есть.
            cur.execute(f"""delete from Grades
where student = (select id from Students where name = '{name}') 
and date = '{date}' and 
subject = (select id from Subjects where name = '{subject}')""")
            con.commit()
            # Потом задаётся новое значение оценки
            cur.execute(f"""insert into Grades(student, subject, Date, value)
Values({ind},{ind2}, '{date}', {grade})""")
            # update работает не во всех случаях,
            # поэтому был выбран именно этот способ
            con.commit()
        except GradeError:
            QMessageBox.question(self, "Ошибка!",
                                 "В нашей школе 5-бальная система оценивания!",
                                 QMessageBox.Ok)
            self.flag1 = False
            self.Ta.setItem(y, x, QTableWidgetItem(""))
            self.flag1 = True
        con.close()

    def change2(self):
        if not self.flag2:
            return
        x, y = [(i.column(), i.row()) for i in self.Ta2.selectedItems()][0]
        try:
            con = sqlite3.connect("Study.db")
            sub = self.Subjects2.currentText()
            period = self.Ta2.horizontalHeaderItem(x).text()
            name = self.Ta2.verticalHeaderItem(y).text()
            value = self.Ta2.item(y, x)
            cur = con.cursor()
            ind_n = cur.execute(f"select id from Students where name = '{name}'").fetchone()[0]
            ind_s = cur.execute(f"select id from Subjects where name = '{sub}'").fetchone()[0]
            if not value or value.text() == "":
                cur.execute(f"""delete from GradeR where
period = '{period}' and student = {ind_n} and subject = {ind_s}""")
                con.commit()
                self.flag2 = False
                self.Ta2.setItem(y, x, QTableWidgetItem(""))
                self.flag2 = True
                return
            value = int(value.text())
            if value < 0 or value > 5:
                raise GradeError
            cur.execute(f"""insert into GradeR(student, subject, period, value)
Values({ind_n}, {ind_s}, '{period}', '{value}')""")
            con.commit()
            con.close()
        except GradeError:
            QMessageBox.question(self, "Ошибка!",
                                 "В нашей школе 5-бальная система оценивания!",
                                 QMessageBox.Ok)
            self.flag2 = False
            self.Ta2.setItem(y, x, QTableWidgetItem(""))
            self.flag2 = True
        self.run2()

    def change3(self):
        if not self.flag3:
            return
        x, y = [(i.column(), i.row()) for i in self.Ta3.selectedItems()][0]
        try:
            name = self.Ta3.verticalHeaderItem(y).text()
            par = self.Ta3.horizontalHeaderItem(x).text()
            params = {"Гендер": "gender", "Дата рождения": "birthday", "Адрес": "adres"}
            par = params[par]
            if not self.Ta3.item(y, x):
                raise KeyError
            value = self.Ta3.item(y, x).text().strip()
            if par == "gender" and value not in ["Мужской", "Женский"]:
                raise NameError
            elif par == "birthday" and not cor_birthday(value):
                raise DateError
            con = sqlite3.connect("Study.db")
            cur = con.cursor()
            cur.execute(f"""Update Students 
set {par} = '{value}' where name = '{name}'""")
            con.commit()
            con.close()
        except NameError:
            QMessageBox.question(self, "Ошибка!", """Неверный гендер. 
Доступные гендеры: Мужской, Женский""", QMessageBox.Ok)
        except DateError:
            QMessageBox.question(self, "Ошибка!", """Неверный формат даты
Допустимые форматы: Год-месяц-день или День-месяц-год.
К примеру: 2222-22-22 или 22-22-2222""", QMessageBox.Ok)
        except KeyError:
            QMessageBox.question(self, "Ошибка!",
                                 "Нельзя вводить пустые строки!", QMessageBox.Ok)
        self.run3()

    def change4(self):
        if not self.flag4:
            return
        try:
            x, y = [(i.column(), i.row()) for i in self.Ta4.selectedItems()][0]
            name = self.Ta4.verticalHeaderItem(y).text()
            par = self.Ta4.horizontalHeaderItem(x).text()
            params = {"Название олимпиады": "olymp", "Баллы": "score", "Занятое место": "res"}
            par = params[par]
            if not self.Ta4.item(y, x):
                raise KeyError
            value = self.Ta4.item(y, x).text()
            con = sqlite3.connect("Study.db")
            cur = con.cursor()
            if not cur.execute(f"""select * from OlympR
where student = (select id from Students where name = '{name}')""").fetchall():
                raise NameError
            elif par == "olymp" and not cur.execute(f"""select name from
             Olymps where name = '{value}'""").fetchone():
                raise OlympError
            elif par == "score" and (500 <= int(value) or int(value) <= 0):
                raise ValueError
            if par != "olymp":
                ol = self.Ta4.item(y, 0).text()
                cur.execute(f"""Update OlympR
set {par} = '{value}' where olymp = (select id from Olymps where name = '{ol}')
and student = (select id from Students where name = '{name}')""")
                con.commit()
            else:
                cur.execute(f"""Update OlympR
set {par} = (select id from Olymps where name = '{value}')
where student = (select id from Students where name = '{name}')""")
                con.commit()
            con.close()
        except KeyError:
            QMessageBox.question(self, "Ошибка!",
                                 "Нельзя вводить пустые строки!", QMessageBox.Ok)
        except NameError:
            QMessageBox.question(self, "Ошибка!",
                                 "Такого ученика нет!", QMessageBox.Ok)
        except OlympError:
            QMessageBox.question(self, "Ошибка!",
                                 "Такой олимпиады нет!", QMessageBox.Ok)
        except ValueError:
            QMessageBox.question(self, "Ошибка!",
                                 "Неверное количество баллов!\n"
                                 "Максимальное количество баллов - 500,"
                                 "\nМинимальное количество баллов - 0.", QMessageBox.Ok)
        self.run4()

    def change5(self):
        if not self.flag5:
            return
        cl = self.Class5.currentText()
        x, y = [(i.column(), i.row()) for i in self.Ta5.selectedItems()][0]
        weekday = self.Ta5.horizontalHeaderItem(x).text()
        weeks = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
        weekday = weeks.index(weekday)
        num = int(self.Ta5.verticalHeaderItem(y).text()[0])
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        res = ["".join(list(i)) for i in cur.execute(f"""select name from Subjects
where min <= {cl} and max >= {cl}""").fetchall()]
        sub, ok = QInputDialog.getItem(self, "Выбор", "Выберите предмет", res)
        if not ok:
            return
        self.flag5 = False
        self.Ta5.setItem(y, x, QTableWidgetItem(sub))
        self.flag5 = True
        ind1 = cur.execute(f"""select id 
from Subjects where name = '{sub}'""").fetchone()[0]
        ind2 = cur.execute(f"""select id 
from Classes where name = '{cl} класс'""").fetchone()[0]
        # Тут надо доработать удаление оценок
        #        ind3 = cur.execute(f"""select subject
        # from Schedule where class = {ind2} and weekday = {weekday} and num = {num}""").fetchone()[0]
        #        cur.execute(f"""delete from Grades where
        # subject = {ind3} and week = {weekday}""")
        cur.execute(f"""delete from Schedule where
class = {ind2} and weekday = {weekday} and num = {num}""")
        con.commit()
        cur.execute(f"""insert into Schedule(subject, class, weekday, num)
Values({ind1},{ind2},{weekday},{num})""")
        con.commit()
        con.close()
        self.run1()

    def class1_change(self):
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = int(self.Class1.currentText())
        res = cur.execute("select name from Subjects "
                          f"where min <= {cl} and max >= {cl}").fetchall()
        self.Subjects1.clear()
        self.Subjects1.addItems([" ".join(list(i)) for i in res])
        con.close()
        self.run1()

    def class2_change(self):
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = int(self.Class2.currentText())
        res = cur.execute("select name from Subjects "
                          f"where min <= {cl} and max >= {cl}").fetchall()
        self.Subjects2.clear()
        self.Subjects2.addItems([" ".join(list(i)) for i in res])
        con.close()
        self.run2()

    def class4_change(self):
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = int(self.Class4.currentText())
        res = cur.execute(f"""select name from Students
where class = (select id from Classes where name = '{cl} класс')""").fetchall()
        self.Students.clear()
        self.Students.addItems(["".join(list(i)) for i in res])
        con.close()
        self.run4()

    def olymps_change(self):
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        res = cur.execute("select name from Olymps").fetchall()
        self.Olymps2.clear()
        self.Olymps2.addItems([" ".join(list(i)) for i in res])
        self.Olymps.clear()
        self.Olymps.addItems([" ".join(list(i)) for i in res] + [""])
        con.close()

    def append_student(self):
        # Эта функция отвечает за добавление учеников
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите добавить ученика?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        try:
            con = sqlite3.connect("Study.db")
            cur = con.cursor()
            name = self.Line.text().strip()
            res = cur.execute(f"select name from Students").fetchall()
            students = [" ".join(list(i)) for i in res]
            if name in students or name == "":
                raise KeyError
            cl = self.Class3.currentText()
            ind = cur.execute(f"""select id from Classes
    where name = '{cl} класс'""").fetchone()[0]
            birth = self.Line3.text()
            if len(birth) != 10:
                raise DateError
            elif not cor_birthday(birth):
                raise DateError
            adres = self.Line2.text()
            if adres == "":
                raise NameError
            gender = self.RaB1.text() if self.RaB1.isChecked() else self.RaB2.text()
            cur.execute(f"""insert into Students(name, class, adres, birthday, gender)
    Values('{name}', {ind}, '{adres}', '{birth}', '{gender}')""")
            con.commit()
            con.close()
        except KeyError:
            QMessageBox.question(self, "Ошибка!",
                                 "Ученик с этим ФИО уже есть!",
                                 QMessageBox.Ok)
        except DateError:
            QMessageBox.question(self, "Ошибка!", """Неверный формат даты
    Допустимые форматы:
    Год-месяц-день или День-месяц-год.
    К примеру: 2222-22-22 или 22-22-2222""", QMessageBox.Ok)
        except NameError:
            QMessageBox.question(self, "Ошибка!",
                                 "Сначала укажите адрес!",
                                 QMessageBox.Ok)
        self.run3()

    def delete_student(self):
        # Эта функция отвечает за удаление учеников
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите исключить ученика?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        try:
            con = sqlite3.connect("Study.db")
            cur = con.cursor()
            cl = self.Class3.currentText()
            res = ["".join(list(i)) for i in cur.execute(f"""select name from Students
where class = (select id from Classes where name = '{cl} класс')""").fetchall()]
            name, ok = QInputDialog.getItem(self, "Выбор", "Выберите ученика", res)
            if not ok:
                return
            if not list(cur.execute(f"""select name 
from Students where name = '{name}'""").fetchone()):
                raise NameError
            ind = cur.execute(f"""select id 
from Students where name = '{name}'""").fetchone()[0]
            cur.execute(f"delete from OlympR where student = {ind}")
            cur.execute(f"delete from Grades where student = {ind}")
            cur.execute(f"delete from Students where name = '{name}'")
            con.commit()
            con.close()
        except NameError:
            QMessageBox.question(self, "Ошибка!",
                                 "Ученика с этим ФИО нет!",
                                 QMessageBox.Ok)
        self.run3()

    def append_olymp(self):
        # Эта функция отвечает за добавление олимпиад в базу данных
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите добавить олимпиаду?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        con = sqlite3.connect("Study.db")
        name, ok = QInputDialog.getText(self, "Введите", "Введите название олимипиады")
        if not ok:
            return
        try:
            if not name:
                raise ValueError
            cur = con.cursor()
            if cur.execute(f"""select name 
from Olymps where name = '{name}'""").fetchone():
                raise OlympError
            cur.execute(f"insert into Olymps(name) Values('{name}')")
            self.Olymps.addItem(name)
            con.commit()
        except OlympError:
            QMessageBox.question(self, "Ошибка!",
                                 "Такая олимпиады уже есть!",
                                 QMessageBox.Ok)
        except ValueError:
            QMessageBox.question(self, "Ошибка!",
                                 "Сначала введите название олимпиады!",
                                 QMessageBox.Ok)
        con.close()
        self.run4()

    def delete_olymp(self):
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите удалить олимпиду?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        try:
            con = sqlite3.connect("Study.db")
            cur = con.cursor()
            res = ["".join(list(i)) for i in cur.execute(f"""select name
from Olymps""").fetchall()]
            ol, ok = QInputDialog.getItem(self, "Выбор", "Выберите олимпиаду", res)
            if not ok:
                return
            cur.execute(f"""delete from OlympR where
olymp = (select id from Olymps where name = '{ol}')""")
            con.commit()
            cur.execute(f"delete from Olymps where name = '{ol}'")
            con.commit()
            con.close()
            self.olymps_change()
            self.run4()
        except OlympError:
            QMessageBox.question(self, "Ошибка!",
                                 "Сначала ведите название олимпиады!",
                                 QMessageBox.Ok)

    def append_olymp_r(self):
        # Эта функция отвечает за добавление в базу данных
        # результатов учеников, участвующих в различных олимпиадах.
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите добавить результат ученика ученика?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        con = sqlite3.connect("Study.db")
        olymp = self.Olymps2.currentText()
        student = self.Students.currentText()
        balls = str(self.BallC.value())
        res = self.Place.text().strip() if self.Place.text().strip() != "" else "100"
        try:
            if not olymp:
                raise ValueError
            cur = con.cursor()
            if not cur.execute(f"""select name
from Olymps where name = '{olymp}'""").fetchone():
                raise OlympError
            if not cur.execute(f"""select name 
from Students where name = '{student}'""").fetchone():
                raise NameError
            ind_o = cur.execute(f"""select id
from Olymps where name = '{olymp}'""").fetchone()[0]
            ind_s = cur.execute(f"""select id
from Students where name = '{student}'""").fetchone()[0]
            cur.execute(f"""insert into OlympR(student, olymp, score, res) 
Values('{ind_s}', '{ind_o}', '{balls}', '{res}')""")
            con.commit()
        except ValueError:
            QMessageBox.question(self, "Ошибка!",
                                 "Сначала введите название олимпиады!",
                                 QMessageBox.Ok)
        except OlympError:
            QMessageBox.question(self, "Ошибка!",
                                 "Такой олимпиады нет!",
                                 QMessageBox.Ok)
        except NameError:
            QMessageBox.question(self, "Ошибка!",
                                 "Такого ученика нет!",
                                 QMessageBox.Ok)
        con.close()
        self.run4()

    def delete_olymp_r(self):
        reply = QMessageBox.question(self, "Подтверждение",
                                     "Вы уверены что хотите удалить результат ученика?",
                                     QMessageBox.Ok | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            return
        ol = self.Olymps2.currentText()
        name = self.Students.currentText()
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cur.execute(f"""delete from OlympR where
olymp = (select id from Olymps where name = '{ol}') 
and student = (select id from Students where name = '{name}')""")
        con.commit()
        con.close()
        self.run4()

    def printe(self):
        # Эта функция отвечает за распечатывание оценок
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        res = cur.execute("select distinct date from Grades").fetchall()
        dates = ["".join(list(i)) for i in res]
        cl = self.Class1.currentText()
        names = [list(i) for i in cur.execute(f"""select name 
from Students where 
class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        grades = [[self.Ta.item(i, j).text() if self.Ta.item(i, j) else "0"
                   for j in range(self.Ta.columnCount())]
                  for i in range(self.Ta.rowCount())]
        sub = self.Subjects1.currentText()
        with open("Grades.txt", 'w', newline='', encoding="utf8") as file:
            file.write(f"Текущие оценки по предмету {sub} "
                       f"учащихся {cl} класса\n")
            file.write(f"ФИО, {', '.join(dates)}, Итоговая\n")
            for i in range(len(names)):
                nam2 = ''.join(names[i])
                gra2 = ', '.join(grades[i][:-1])
                res = grades[i][-1]
                file.write(f"{nam2}\t{gra2}\t{res}\n")
        con.close()
        self.preview = Preview("Grades.txt")
        self.preview.show()

    def printe2(self):
        # Эта функция отвечает за распечатывание итоговых оценок
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class2.currentText()
        sub = self.Subjects2.currentText()
        students = ["".join(list(i)) for i in cur.execute(f"""select name
from Students where class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        A = ["1 четверть", "2 четверть", "3 четверть", "4 четверть", "Годовая"]
        grades = [[self.Ta2.item(i, j).text() if self.Ta.item(i, j) else "0"
                   for j in range(self.Ta2.columnCount())]
                  for i in range(self.Ta2.rowCount())]
        with open("GradeR.txt", 'w', newline='', encoding="utf8") as file:
            file.write(f"Итоговые оценки по предмету {sub} "
                       f"учащихся {cl} класса\n")
            file.write(f"ФИО, {', '.join(A)}\n")
            for i in range(len(students)):
                name = students[i]
                file.write(f"{name}\t{','.join(grades[i])}\n")
        con.close()
        self.preview = Preview("GradeR.txt")
        self.preview.show()

    def printe3(self):
        # Эта функция отвечает за распечатывание данных об
        # учениках, их гендере, дате рождения и адресе проживания
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class3.currentText()
        res = [list(i) for i in cur.execute(f"""select name, gender, birthday, adres
from Students where class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        with open("Students.txt", 'w', newline='', encoding="utf8") as file:
            file.write(f"Учащиеся {cl} класса\n")
            file.write(f"ФИО, Гендер, Дата рождения, Адрес проживания\n")
            for i in range(len(res)):
                file.write(f"{', '.join(res[i])}\n")
        con.close()
        self.preview = Preview("Students.txt")
        self.preview.show()

    def printe4(self):
        # Эта функция отвечает за распечатывание результатов олимпиад
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class4.currentText()
        ol = self.Olymps.currentText()
        if ol == "":
            res = [list(map(str, i)) for i in cur.execute(f"""select Students.name, Olymps.name,
OlympR.score, OlympR.res from OlympR
inner join Olymps on OlympR.olymp=Olymps.id
inner join Students on OlympR.student=Students.id where
Students.class = (select id from Classes where name = '{cl} класс')""").fetchall()]
            with open("./Отчёты/Olymp results.txt", 'w', newline='', encoding="utf8") as file:
                file.write("Результаты учащихся во всех олимпиадах\n")
                file.write(f"Название олимпиады, ФИО участника, набранные баллы, занятое место\n")
                for i in range(len(res)):
                    file.write(f"{', '.join(res[i])}\n")
        else:
            res = [list(map(str, i)) for i in cur.execute(f"""select Students.name, 
OlympR.score, OlympR.res from OlympR
inner join Olymps on OlympR.olymp=Olymps.id
inner join Students on OlympR.student=Students.id
where Olymps.name = '{ol}' and
Students.class = (select id from Classes where name = '{cl} класс')""").fetchall()]
            with open("Olymp results.txt", 'w', newline='', encoding="utf8") as file:
                file.write(f"Результаты учащихся в олимпиаде '{ol}'\n")
                file.write("ФИО участника, набранные баллы, занятое место\n")
                for i in range(len(res)):
                    file.write(f"{', '.join(res[i])}\n")
        con.close()
        self.preview = Preview("Olymp results.txt")
        self.preview.show()

    def run1(self):
        # Эта функция ответственена за создание таблицы с оценками
        self.flag1 = False
        self.Ta.clear()
        subject = self.Subjects1.currentText()
        cl = self.Class1.currentText()
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        # Сперва задаются заголовки с датами и именами
        weeks = [i[0] for i in cur.execute(f"""select weekday from Schedule
where subject=(select id from Subjects where name = '{subject}')
and class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        ch = int(self.Period.currentText()[0])
        dates = []
        for i in range(self.chetvs[ch][1]):
            date = self.chetvs[ch][0] + dt.timedelta(days=i)
            if date.weekday() in weeks:
                dates.append(date.strftime("%b-%d"))
        self.Ta.setColumnCount(len(dates) + 1)
        self.Ta.setHorizontalHeaderLabels(dates + ["Итог"])
        res = cur.execute(f"""select distinct Students.name from Students
inner join Classes on Students.class=Classes.id
where Classes.name = '{cl} класс'""").fetchall()
        students = ["".join(i) for i in res]
        students.sort()
        self.Ta.setRowCount(len(students))
        self.Ta.setVerticalHeaderLabels(students)
        # Затем таблица заполняется оценками
        res = cur.execute(f"""select Students.name, Grades.date, Grades.value 
from Grades
inner join Students on Grades.student=Students.id
inner join Subjects on Grades.subject=Subjects.id
inner join Classes on Students.class=Classes.id
where Subjects.name = '{subject}' and Classes.name = '{cl} класс'""").fetchall()
        grades = [list(i) for i in res]
        for i in grades:
            ind1 = students.index(i[0])
            if i[1] not in dates:
                continue
            ind2 = dates.index(i[1])
            grade = i[2]
            self.Ta.setItem(ind1, ind2, QTableWidgetItem(grade))
        # Затем выстовляются итоговые оценки за четверть
        ind2 = self.Ta.columnCount() - 1
        for i in range(self.Ta.rowCount()):
            grades2 = []
            for j in range(self.Ta.columnCount() - 1):
                if self.Ta.item(i, j) and self.Ta.item(i, j).text().isdigit():
                    grades2.append(int(self.Ta.item(i, j).text()))
            grades2 = str(round(sum(grades2) / len(grades2), 2)) if grades2 else "0"
            self.Ta.setItem(i, ind2, QTableWidgetItem(grades2))
        con.close()
        # Этот флаг нужен чтобы работала другая функция
        self.flag1 = True

    def run2(self):
        self.flag2 = False
        self.Ta2.clear()
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class2.currentText()
        sub = self.Subjects2.currentText()
        res = ["".join(list(i)) for i in cur.execute(f"""select name from Students
where class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        self.Ta2.setRowCount(len(res))
        self.Ta2.setVerticalHeaderLabels(res)
        A = ["1 четверть", "2 четверть", "3 четверть", "4 четверть", "Годовая"]
        self.Ta2.setColumnCount(len(A))
        self.Ta2.setHorizontalHeaderLabels(A)
        students = ["".join(list(i)) for i in cur.execute(f"""select 
name from Students where
class = (select id from Classes where name = '{cl} класс')""").fetchall()]
        for j in range(5):
            res = [list(i) for i in cur.execute(f"""select 
Students.name, GradeR.value from GradeR
inner join Students on GradeR.student=Students.id
where GradeR.subject = (select id from Subjects where name = '{sub}') and
Students.class = (select id from Classes where name = '{cl} класс') and
GradeR.period = '{A[j]}'""").fetchall()]
            for i in range(len(res)):
                ind = students.index(res[i][0])
                self.Ta2.setItem(ind, j, QTableWidgetItem(res[i][1]))
        con.close()
        self.flag2 = True

    def run3(self):
        # Эта функция отвечает за создание списка учеников
        self.flag3 = False
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class3.currentText()
        res = cur.execute(f"""select name, gender, birthday, adres from Students 
    where class = (select id from Classes where name = '{cl} класс')""").fetchall()
        res = [list(i) for i in res]
        res2 = ["Гендер", "Дата рождения", "Адрес"]
        self.Ta3.setColumnCount(3)
        self.Ta3.setHorizontalHeaderLabels(res2)
        self.Ta3.setRowCount(len(res))
        self.Ta3.setVerticalHeaderLabels([i[0] for i in res])
        for i in range(len(res)):
            self.Ta3.setItem(i, 0, QTableWidgetItem(res[i][1]))
            self.Ta3.setItem(i, 1, QTableWidgetItem(res[i][2]))
            self.Ta3.setItem(i, 2, QTableWidgetItem(res[i][3]))
        con.close()
        self.flag3 = True

    def run4(self):
        # Эта функция отвечает за создание таблицы с результами олимпиад
        self.flag4 = False
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        cl = self.Class4.currentText()
        if self.Olymps.currentText() != "":
            ol = self.Olymps.currentText()
        else:
            ol = "%"
        res = cur.execute(f"""select Students.name, Olymps.name, 
OlympR.score, OlympR.res from OlympR
inner join Students on OlympR.student=Students.id
inner join Classes on Students.class=Classes.id
inner join Olymps on OlympR.olymp=Olymps.id
where Classes.name = '{cl} класс' and Olymps.name like '{ol}'""").fetchall()
        res = [list(i) for i in res]
        self.Ta4.setRowCount(len(res))
        self.Ta4.setVerticalHeaderLabels([i[0] for i in res])
        for i in range(len(res)):
            self.Ta4.setItem(i, 0, QTableWidgetItem(res[i][1]))
            self.Ta4.setItem(i, 1, QTableWidgetItem(res[i][2]))
            self.Ta4.setItem(i, 2, QTableWidgetItem(res[i][3]))
        con.close()
        self.flag4 = True

    def run5(self):
        self.flag5 = False
        cl = self.Class5.currentText()
        con = sqlite3.connect("Study.db")
        cur = con.cursor()
        self.Ta5.clear()
        self.Ta5.setRowCount(7)
        self.Ta5.setVerticalHeaderLabels([f"{i + 1} урок" for i in range(7)])
        self.Ta5.setColumnCount(6)
        weeks = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
        self.Ta5.setHorizontalHeaderLabels(weeks)
        for i in range(6):
            res = [list(i) for i in cur.execute(f"""select
Subjects.name, Schedule.num from Schedule
inner join Subjects on Schedule.subject=Subjects.id
where Schedule.weekday = {i}
and Schedule.class = (select id from Classes where name = '{cl} класс')
order by Schedule.num""").fetchall()]
            for j in range(len(res)):
                self.Ta5.setItem(res[j][1] - 1, i, QTableWidgetItem(res[j][0]))
        con.close()
        self.flag5 = True


class Preview(QWidget):
    def __init__(self, file):
        super().__init__()
        uic.loadUi("interface2.ui", self)
        self.file = file
        self.Yes.clicked.connect(self.print)
        self.No.clicked.connect(self.close)
        with open(file, "r", encoding="utf8") as f:
            self.TeB.setText(f.read())

    def print(self):
        os.startfile("Grades.txt", "print")
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec_())
