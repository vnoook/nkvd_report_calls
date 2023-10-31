import os
import sys
import openpyxl
import PyQt5
import PyQt5.QtWidgets


# import PyQt5.QtCore
# import PyQt5.QtGui


# класс главного окна
class WindowMain(PyQt5.QtWidgets.QMainWindow):
    """Класс главного окна"""

    # описание главного окна
    def __init__(self):
        super().__init__()

        # переменные
        self.info_extention_open_file_xlsx = 'Файлы XLSX (*.xlsx)'
        self.info_path_open_file = None
        self.text_empty_path_file = 'файл пока не выбран'
        self.info_for_open_file = 'Выберите XLSX файл (.XLSX)'

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Парсер XLSX файлов для отчёта Журнал записей пациентов')
        self.setGeometry(450, 100, 700, 180)
        self.setWindowFlags(PyQt5.QtCore.Qt.WindowStaysOnTopHint)

        # ОБЪЕКТЫ НА ФОРМЕ
        # label_select_file
        self.label_select_file = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file.setObjectName('label_select_file')
        self.label_select_file.setText('Выберите файл XLSX')
        self.label_select_file.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file.setFont(font)
        self.label_select_file.adjustSize()
        self.label_select_file.setToolTip(self.label_select_file.objectName())

        # toolButton_select_file_xlsx
        self.toolButton_select_file_xlsx = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_xlsx.setObjectName('toolButton_select_file_xlsx')
        self.toolButton_select_file_xlsx.setText('...')
        self.toolButton_select_file_xlsx.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_file_xlsx.setFixedWidth(50)
        self.toolButton_select_file_xlsx.clicked.connect(self.select_file_xlsx)
        self.toolButton_select_file_xlsx.setToolTip(self.toolButton_select_file_xlsx.objectName())

        # label_path_file
        self.label_path_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file.setObjectName('label_path_file')
        self.label_path_file.setEnabled(False)
        self.label_path_file.setText(self.text_empty_path_file)
        self.label_path_file.setGeometry(PyQt5.QtCore.QRect(10, 70, 400, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_path_file.setFont(font)
        self.label_path_file.adjustSize()
        self.label_path_file.setToolTip(self.label_path_file.objectName())

        # pushButton_parse_to_xls
        self.pushButton_parse_to_xls = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_parse_to_xls.setObjectName('pushButton_parse_to_xls')
        self.pushButton_parse_to_xls.setEnabled(False)
        self.pushButton_parse_to_xls.setText('Создать отчёт "Журнал записей пациентов"')
        self.pushButton_parse_to_xls.setGeometry(PyQt5.QtCore.QRect(10, 100, 260, 25))
        self.pushButton_parse_to_xls.clicked.connect(self.parse_xlsx)
        self.pushButton_parse_to_xls.setToolTip(self.pushButton_parse_to_xls.objectName())

        # EXIT
        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 140, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file_xlsx(self):
        # переменная для хранения информации из окна выбора файла
        data_of_open_file_name = None

        # запоминание старого значения пути выбора файлов
        old_path_of_selected_xlsx_file = self.label_path_file.text()

        # непосредственное окно выбора файла и переменная для хранения пути файла
        data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                             self.info_for_open_file,
                                                                             self.info_path_open_file,
                                                                             self.info_extention_open_file_xlsx)

        # выбор только пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]

        # нажата кнопка выбора XLSX файла
        if file_name == '':
            self.label_path_file.setText(old_path_of_selected_xlsx_file)
            self.label_path_file.adjustSize()
        else:
            old_path_of_selected_xlsx_file = self.label_path_file.text()
            self.label_path_file.setText(file_name)
            self.label_path_file.adjustSize()

        # активация и деактивация объектов на форме зависящее от выбора файла
        if self.text_empty_path_file not in self.label_path_file.text():
            self.pushButton_parse_to_xls.setEnabled(True)

    # функция создания отчёта
    def parse_xlsx(self):
        # получение пути и имени выбранного файла
        file_xlsx = self.label_path_file.text()
        file_xlsx_path = os.path.split(file_xlsx)[0]
        file_xlsx_name = os.path.split(file_xlsx)[1]

        # открывается выбранный файл
        wb_in = openpyxl.load_workbook(file_xlsx)
        wb_in_s = wb_in['Журнал записей пациентов']

        # строка начала и конца
        wb_in_s_row_begin = 3
        wb_in_s_row_end = wb_in_s.max_row - 1

        # номера колонок для сбора данных
        wb_in_s_col_1 = 7
        wb_in_s_col_2 = 18
        wb_in_s_col_3 = 19

        # структуры для сбора данных
        list_all_data = []
        # set_col_1 = set()  # отделение
        # set_col_2 = set()  # записавшая организация
        # set_col_3 = set()  # кем записан

        # словарь для хранения отделений
        dict_departments = {}
        # словарь для хранения записавших организаций
        dict_organization = {}
        # словарь для хранения "кеи записан"
        dict_persona = {}

        # строки, которые нужно складывать
        str_for_summ = {'Интеграция Е.Р.': 'ЕПГУ-Госуслуги',
                        'Administrator A.A.': 'МИАЦ',
                        'Система Г.С.': 'СГС-Робот Николай'}
        # строки по которым нужно фильтровать
        str_for_filter = ('Амбулаторное отделение №1', 'Амбулаторное отделение №2',
                          'Амбулаторное отделение №3', 'Амбулаторное отделение №4',
                          'Подростковый специализированный центр профилактики и лечения инфекций, передаваемых половым путем')

        # получение всех данных из файла и его закрытие, чтобы к нему больше не возвращаться
        for row in range(wb_in_s_row_begin, wb_in_s_row_end + 1):
            list_all_data.append([wb_in_s.cell(row=row, column=wb_in_s_col_1).value,
                                  wb_in_s.cell(row=row, column=wb_in_s_col_2).value,
                                  wb_in_s.cell(row=row, column=wb_in_s_col_3).value
                                  ])
        wb_in.close()

        # создание отчёта в xlsx и активация рабочего листа
        wb_out = openpyxl.Workbook()
        wb_out_s = wb_out.active

        # подсчёт и распределение
        for val_str in list_all_data:
            if val_str[0] in str_for_filter:
                # заполнение словаря отделений
                if dict_departments.get(val_str[0]) is None:
                    dict_departments[val_str[0]] = 1
                else:
                    dict_departments[val_str[0]] = dict_departments[val_str[0]] + 1

                # заполнение словаря записавших организаций
                if dict_organization.get(val_str[0]) is None:
                    dict_organization[val_str[0]] = {val_str[1]: 1}
                else:
                    if dict_organization[val_str[0]].get(val_str[1]) is None:
                        dict_organization[val_str[0]][val_str[1]] = 1
                    else:
                        dict_organization[val_str[0]][val_str[1]] = dict_organization[val_str[0]][val_str[1]] + 1

                # заполнение словаря "кем записан"
                if val_str[2] in str_for_summ:
                    if dict_persona.get(val_str[0]) is None:
                        dict_persona[val_str[0]] = {val_str[2] : 1}
                    else:
                        if dict_persona[val_str[0]].get(val_str[2]) is None:
                            dict_persona[val_str[0]][val_str[2]] = 1
                        else:
                            dict_persona[val_str[0]][val_str[2]] = dict_persona[val_str[0]][val_str[2]] + 1

        # print(dict_persona)
        # TODO
        # сортировка словарей по ключам

        # формирование отчёта
        for k, v in dict_organization.items():
            print(k, '...', v)

            # print(f'{k} - {dict_departments[k]} пациент(ов)')
            wb_out_s.append([f'{k} - {dict_departments[k]} пациент(ов)'])

            # print('(--- строка подсчёта по которой есть вопросы ---)')
            wb_out_s.append(['(--- строка подсчёта по которой есть вопросы ---)'])

            for d, q in v.items():
                # print(d, '-', q)
                wb_out_s.append([f'{d} - {q}'])

            # print()
            wb_out_s.append([])

        # создание названия выходного файла xls
        file_xls_path = file_xlsx_path[:]
        file_xls_name = os.path.splitext(file_xlsx_name)[0] + '_отчёт.xlsx'
        file_report = os.path.abspath(os.path.join(file_xls_path, file_xls_name))

        # сохранение файла xlsx и закрытие его
        wb_out.save(file_report)
        wb_out.close()

        # открытие папки с сохранённым файлом xls
        fullpath = os.path.abspath(file_xls_path)
        PyQt5.QtGui.QDesktopServices.openUrl(PyQt5.QtCore.QUrl.fromLocalFile(fullpath))

    # событие - нажатие на кнопку Выход
    @staticmethod
    def click_on_btn_exit():
        sys.exit()


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app_window_main = WindowMain()
    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()
