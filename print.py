import os  # Отсюда нам понадобятся методы для отображения содержимого директорий
import sys  # sys нужен для передачи argv в QApplication
from pathlib import Path

from PyQt5 import QtWidgets

from qt import mainwindow  # Это наш конвертированный файл дизайна


class ExampleApp(QtWidgets.QMainWindow, mainwindow.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.btnBrowse.clicked.connect(
            self.browse_folder
        )  # Выполнить функцию browse_folder
        # при нажатии кнопки

    def browse_folder(self):
        self.listWidget.clear()  # На случай, если в списке уже есть элементы
        dialog = QtWidgets.QFileDialog(self)
        (filename, _) = dialog.getOpenFileName(
            dialog, "Выберите таблицу", filter="Таблицы (*.xls *.xlsx)"
        )

        if not filename:
            return
        file = Path(filename)

        self.buttonBox.setEnabled(True)
        self.setWindowTitle(f'Active - {file.name}')


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
