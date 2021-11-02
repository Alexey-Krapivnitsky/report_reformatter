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
        self.aup_data = {}
        self.orion_data = []
        self.setFixedSize(530, 200)
        self.setWindowTitle('by Lekks. Слияние отчетов СКУД и АУП')
        self.orion_btn = QtWidgets.QPushButton('Выбрать файл отчета СКУД')
        self.orion_btn.setFixedSize(200, 30)
        self.aup_btn = QtWidgets.QPushButton('Выбрать файл АУП')
        self.aup_btn.setFixedSize(200, 30)
        self.orion_path = QtWidgets.QLineEdit()
        self.orion_path.setFixedSize(300, 30)
        self.orion_path.setEnabled(False)
        self.aup_path = QtWidgets.QLineEdit()
        self.aup_path.setFixedSize(300, 30)
        self.aup_path.setEnabled(False)
        self.start_btn = QtWidgets.QPushButton('Слияние')
        self.start_btn.setEnabled(False)
        self.orion_box = QtWidgets.QHBoxLayout()
        self.aup_box = QtWidgets.QHBoxLayout()
        self.main_box = QtWidgets.QVBoxLayout()
        self.orion_box.addWidget(self.orion_btn)
        self.orion_box.addWidget(self.orion_path)
        self.aup_box.addWidget(self.aup_btn)
        self.aup_box.addWidget(self.aup_path)
        self.main_box.addLayout(self.orion_box)
        self.main_box.addLayout(self.aup_box)
        self.main_box.addWidget(self.start_btn)
        self.central = QtWidgets.QWidget()
        self.central.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.central.setLayout(self.main_box)
        self.setCentralWidget(self.central)
        self.orion_btn.clicked.connect(self.orion_file_change)
        self.aup_btn.clicked.connect(self.aup_file_change)
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

    def aup_file_change(self):
        self.aup_file = QtWidgets.QFileDialog.getOpenFileName(
            parent=self,
            caption="Выбор файла АУП",
            directory=QtCore.QDir.currentPath(),
            filter="All (*);; ",
            initialFilter="")
        self.aup_path.setText(self.aup_file[0])
        self.aup_orion_is_no_empty()

    def aup_orion_is_no_empty(self):
        if self.orion_path.text() and self.aup_path.text():
            self.start_btn.setEnabled(True)

    def csv_dict_reader(self, file_obj):
        reader = csv.DictReader(file_obj, delimiter=',')

        for line in reader:
            cc = list(line.values())
            kk = list(set(line.keys()))
            self.orion_data.append(cc[0].split(';'))
        self.orion_data.insert(0, kk[0].split(';'))
        for elem in self.orion_data:
            del (elem[2])
            del (elem[3:])
            if elem[0] == 'Сотрудник':
                elem.append('Причины отсутствия')

    def xls_dict_reader(self, file_obj):
        records = pyexcel.get_records(file_name=file_obj)
        pars = []

        for val in records:
            pars.append(dict(val))

        for item in pars:
            if item['-2'] not in self.workers.keys():
                self.workers.setdefault(item['-2'],
                                        [f"{item['-7']}: {item['-5']}/{item['-6']}"]
                                        if item['-6'] else [f"{item['-7']}: {item['-5']}"])
            else:
                self.workers[item['-2']].append(f"{item['-7']}: {item['-5']}/{item['-6']}"
                                                if item['-6'] else f"{item['-7']}: {item['-5']}")

    def merger(self):
        self.start_btn.setEnabled(False)
        with open(self.orion_file[0]) as f_obj:
            self.csv_dict_reader(f_obj)
        self.xls_dict_reader(self.aup_file[0])
        for key in sorted(self.workers.keys()):
            temp_worker_val = []
            for i in range(len(self.workers[key])):
                if self.workers[key][i].split(':')[0] not in ['Я', 'Н', 'ВО', 'ТУ', 'РН', 'РВ', 'Х']:
                    temp_worker_val.append(str(self.workers[key][i]))
            self.workers[key] = ', '.join(temp_worker_val)
            for elem in self.orion_data:
                if key and elem[0] == key:
                    elem.append(self.workers[key])
        new_report = QtWidgets.QFileDialog.getSaveFileName(
            parent=self,
            caption="Сохранение объединенного отчета",
            directory=QtCore.QDir.currentPath(),
            filter="All (*);; Документ Excel (*.xls)",
            initialFilter="Документ Excel (*.xls)")
        pyexcel.save_as(array=self.orion_data, dest_file_name=new_report[0])
        QtWidgets.QMessageBox.information(self, "Слияние файлов", f"Объединенный файл успешно создан,\n"
                                                                  f"новый файл находится в папке:\n"
                                                                  f"{new_report[0]}")
        self.aup_path.clear()
        self.orion_path.clear()
        self.orion_file = ''
        self.aup_file = ''
        self.workers = {}
        self.aup_data = {}
        self.orion_data = []


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MW = FileSelect()
    MW.show()
    sys.exit(app.exec_())
