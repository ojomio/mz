import itertools
import sys  # sys нужен для передачи argv в QApplication
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem

from qt import mainwindow  # Это наш конвертированный файл дизайна


class ExampleApp(QtWidgets.QMainWindow, mainwindow.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.buttonBox.hide()
        self.btnBrowse.clicked.connect(
            self.browse_folder
        )  # Выполнить функцию browse_folder
        # при нажатии кнопки

        self.buttonBox.accepted.connect(self.process_file)
        # self.buttonBox.rejected.connect(self.process_file)

        self.file: Optional[Path] = None

    def browse_folder(self):
        self.buttonBox.hide()
        self.tableWidget.clear()  # На случай, если в списке уже есть элементы
        dialog = QtWidgets.QFileDialog(self)
        (filename, _) = dialog.getOpenFileName(
            dialog, "Выберите таблицу", filter="Таблицы (*.xls *.xlsx)"
        )

        self._handle_browse_result(filename)

    def _handle_browse_result(self, filename: Optional[str]):
        if not filename:
            return

        self.file = Path(filename)

        self.wb = openpyxl.open(self.file)
        self.ws = self.wb.active

        self.buttonBox.setEnabled(True)
        self.setWindowTitle(f'Active - {self.file.name}')
        self._load_table_widget(self.ws)
        self.buttonBox.show()

    def _load_table_widget(self, ws: Worksheet):
        headers = [
            cell.value for cell in itertools.takewhile(lambda x: bool(x), ws['1'])
        ]
        self.tableWidget.setColumnCount(len(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)

        data = ws.iter_rows(min_row=2)
        for x, row in enumerate(data):
            if row[0].value is None:
                break

            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            for y, cell in enumerate(row):
                val = cell.value
                if val is None:
                    continue
                item = QTableWidgetItem(str(val))
                self.tableWidget.setItem(x, y, item)

        self.range_start.setText('A1')
        self.range_end.setText(f'A{ws.max_row}')

    def process_file(self):
        if not self.file:
            return


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
