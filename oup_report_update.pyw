import csv
import pyexcel
import sys
from PyQt5 import QtWidgets, QtCore


class FileSelect(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.orion_file = ''
        self.aup_file = ''
        self.workers = {}
        self.out_data = [['Сотрудник', 'Должность', 'Табельный номер', 'Отдел', 'Режим работы',
                          'Дата', 'Отработал', 'Начало дня', 'Конец дня']]
        self.setFixedSize(530, 200)
        self.setWindowTitle('by Lekks. Реформатирование отчета Intellect')
        self.orion_btn = QtWidgets.QPushButton('Выбрать файл отчета')
        self.orion_btn.setFixedSize(200, 30)
        self.orion_path = QtWidgets.QLineEdit()
        self.orion_path.setFixedSize(300, 30)
        self.orion_path.setEnabled(False)
        self.start_btn = QtWidgets.QPushButton('Начать преобразование')
        self.start_btn.setEnabled(False)
        self.orion_box = QtWidgets.QHBoxLayout()
        self.main_box = QtWidgets.QVBoxLayout()
        self.orion_box.addWidget(self.orion_btn)
        self.orion_box.addWidget(self.orion_path)
        self.main_box.addLayout(self.orion_box)
        self.main_box.addWidget(self.start_btn)
        self.central = QtWidgets.QWidget()
        self.central.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.central.setLayout(self.main_box)
        self.setCentralWidget(self.central)
        self.orion_btn.clicked.connect(self.orion_file_change)
        self.start_btn.clicked.connect(self.merger)

    def orion_file_change(self):
        self.orion_file = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption="Выбор файла СКУД",
            directory=QtCore.QDir.currentPath(),
            filter="All (*);; ",
            initialFilter="")
        self.orion_path.setText(self.orion_file[0])
        self.aup_orion_is_no_empty()

    def aup_orion_is_no_empty(self):
        if self.orion_path.text():
            self.start_btn.setEnabled(True)

    def csv_dict_reader(self, file_obj):

        employee = None
        employee_post = None
        service_number = None
        department = None
        work_schedule = None
        employee_data = {}
        new_employee = False
        report_data = {}

        reader = csv.DictReader(file_obj, delimiter='\n')
        for line in reader:
            data = list(line.values())[0].split(';')
            if data[0].find('Регион') == 0:
                new_employee = True
                employee_data = {}
            elif data[0].find('Вход в регион') == 0:
                new_employee = False

            if new_employee:
                if data[1].find('Табельный номер') == 0:
                    service_number = data[1].split(': ')[1]
                    employee_data.setdefault('Табельный номер', service_number)
                elif data[1].find('Должность') == 0:
                    employee_post = data[1].split(': ')[1]
                    employee_data.setdefault('Должность', employee_post)
                elif data[4].find('Отдел') == 0:
                    department = data[4].split(': ')[1]
                    employee_data.setdefault('Отдел', department)
                elif data[4].find('График работы') == 0:
                    work_schedule = data[4].split(': ')[1]
                    employee = data[1]
                    employee_data.setdefault('График работы', work_schedule)
                    employee_data.setdefault('Сотрудник', employee)
                if len(employee_data) == 5:
                    report_data.setdefault(employee, [[employee_post, service_number, department, work_schedule]])
            if not new_employee and data[0].find('Присутствие') != 0 and data[0].find('Дата') != 0 \
                    and data[0].find('Вход в регион') != 0:
                if data[0] and data[0] != 'Итого по сотруднику:':
                    report_date = data[0].split(' ')[0]
                    in_time = data[0].split(' ')[1]
                    out_time = data[2].split(' ')[1]
                    total_work_time = ':'.join(data[6].split(':')[:2])
                    report_data[employee].append([report_date, total_work_time, in_time, out_time])

        for key, value in report_data.items():
            for elem in value[1:]:
                self.out_data.append([key, *value[0], *elem])

    def merger(self):
        self.start_btn.setEnabled(False)
        with open(self.orion_file[0]) as f_obj:
            self.csv_dict_reader(f_obj)

        new_report = QtWidgets.QFileDialog.getSaveFileName(
            parent=self,
            caption="Сохранение преобразованного отчета",
            directory=QtCore.QDir.currentPath(),
            filter="All (*);; Документ Excel (*.xls)",
            initialFilter="Документ Excel (*.xls)")
        pyexcel.save_as(array=self.out_data, dest_file_name=new_report[0])
        QtWidgets.QMessageBox.information(self, "Преобразование отчета", f"Файл успешно создан,\n"
                                                                         f"новый файл находится в папке:\n"
                                                                         f"{new_report[0]}")
        self.orion_path.clear()
        self.orion_file = ''
        self.out_data = [['Сотрудник', 'Должность', 'Табельный номер', 'Отдел', 'Режим работы',
                          'Дата', 'Отработал', 'Начало дня', 'Конец дня']]


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MW = FileSelect()
    MW.show()
    sys.exit(app.exec_())
