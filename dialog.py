import sys

from PyQt5 import uic, QtGui
from PyQt5.QtWidgets import QDialog, QPushButton, QGridLayout, QLabel, QHBoxLayout
from PyQt5.QtCore import Qt


class QDialogClass(QDialog):
    def __init__(self, parent=None, main=None):
        super(QDialogClass, self).__init__(parent)
        self.parent = parent
        self.main = main
        self.setWindowTitle("HSEReportChecker")
        self.setStyleSheet("QWidget {background-color: #A0AECD;}")
        self.setFixedSize(500, 150)
        self.label = QLabel("Хотите сохранить настройки проверки?")
        self.label.setFont(QtGui.QFont("Verdana", 14, QtGui.QFont.Bold))

        self.layout = QGridLayout()

        self.layout.addWidget(self.label, 0, 0, 1, 3, alignment=Qt.AlignCenter)

        self.save_button = QPushButton("Да")
        self.save_button.clicked.connect(self.save_settings)
        self.save_button.setStyleSheet(
            "QPushButton {min-width: 140px; max-width: 180px; min-height: 50px; color: solid black; font-weight: 600; "
            "border-radius: 8px; border: 1px solid black; outline: 0px; } QPushButton:hover { background-color: "
            "#6E6E6E; border: 1px solid black; }")
        self.save_button.setFont(QtGui.QFont("Verdana", 14, QtGui.QFont.Bold))
        self.layout.addWidget(self.save_button, 1, 0, alignment=Qt.AlignLeft)

        self.discard_button = QPushButton("Нет")
        self.discard_button.clicked.connect(self.accept)
        self.discard_button.setStyleSheet(
            "QPushButton {min-width: 140px; max-width: 180px; min-height: 50px; color: solid black; font-weight: 600; "
            "border-radius: 8px; border: 1px solid black; outline: 0px; } QPushButton:hover { background-color: "
            "#6E6E6E; border: 1px solid black; }")
        self.discard_button.setFont(QtGui.QFont("Verdana", 14, QtGui.QFont.Bold))
        self.layout.addWidget(self.discard_button, 1, 1, alignment=Qt.AlignCenter)

        self.cancel_button = QPushButton("Отмена")
        self.cancel_button.clicked.connect(self.reject)
        self.cancel_button.setStyleSheet(
            "QPushButton {min-width: 140px; max-width: 180px; min-height: 50px; color: solid black; font-weight: 600; "
            "border-radius: 8px; border: 1px solid black; outline: 0px; } QPushButton:hover { background-color: "
            "#6E6E6E; border: 1px solid black; }")
        self.cancel_button.setFont(QtGui.QFont("Verdana", 14, QtGui.QFont.Bold))
        self.layout.addWidget(self.cancel_button, 1, 2, alignment=Qt.AlignRight)

        self.setLayout(self.layout)

    def save_settings(self):
        # Здесь вызывается функция сохранения настроек
        if self.main is None:
            self.parent.saveSettings(self.parent.geometry())
        else:
            self.main.saveSettings(self.parent.geometry())
        self.accept()  # Закрыть окно


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)
