import os
import subprocess
import sys
import pathlib
import multiprocessing
from multiprocessing.managers import SyncManager, BaseManager
import time


from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtCore import QRunnable, Qt, QThreadPool, QThread, QObject, pyqtSignal, QSettings, QRect
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QDialog

from dialog import QDialogClass
from settings import SettingsWindow
import keyboard
from docx.enum.section import WD_ORIENTATION
from docx import Document

from docxParser import DocumentParser
from ui_mainwindow import Ui_Checker


def parse_docx(filename, p):
    file, extension = os.path.splitext(filename)
    if extension == ".docx":
        p.parse_document(filename)


def get_internal_data(filename, *add):
    if os.path.isdir("_internal"):
        return f"_internal\\{filename}"
    if len(add) > 0:
        return f"{"\\".join(add)}\\{filename}"
    return filename


# Step 1: Create a worker class
class Worker(QThread):
    doc_progressed = pyqtSignal()

    def __init__(self, filenames, parser):
        super().__init__()
        self.filenames = filenames
        self.parser = parser

    def doc_finished(self, res):
        self.doc_progressed.emit()

    def run(self):
        pool = multiprocessing.Pool()
        parser = self.parser
        for filename in self.filenames:
            pool.apply_async(parse_docx, args=(filename, parser), callback=self.doc_finished)
        pool.close()
        pool.join()


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_Checker()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon(get_internal_data('icon.png', 'icon')))
        self.SettingsWindow = SettingsWindow(self)
        self.SettingsWindow.setWindowIcon(QtGui.QIcon(get_internal_data('icon.png', 'icon')))
        self.fileFlag = False
        self.fileNames = []  # Список путей до выбранных файлов
        # self.ui = uic.loadUi('uiMainFile.ui', self)  # Открытие файла ui
        # Параметры форматирования документа
        self.text_checklist = None
        self.heading1_checklist = None
        self.heading2_checklist = None
        self.heading3_checklist = None
        self.page_checklist = None
        self.list_checklist = None
        self.table_checklist = None
        self.picture_checklist = None
        self.table_heading_checklist = None
        self.title_picture_checklist = None
        self.table_title_checklist = None
        self.enable_optional_settings = None

        self.setWindowTitle("HSEReportChecker")

        self.dlg = None
        self.settings = QSettings("Checker", "CheckerUserConfigs", self)
        self.loadSettings()

        self.proc = None
        self.explorer = None
        self.file_path = pathlib.Path(__file__).parent.resolve()

        self.ui.LEResLine.hide()
        self.ui.btnOpenRes.hide()

        keyboard.add_hotkey("f1", self.openDocumentation)

        self.ui.LEErrorLine.hide()

        # Привязки функций к кнопкам
        self.ui.btnChooseFile.clicked.connect(self.chooseFile)     # Кнопка выбора файла
        self.ui.btnRun.clicked.connect(self.runCheckingCorrectness)    # Кнопка запуска проверки файла
        self.ui.btnSettings.clicked.connect(self.setupSettings)    # Кнопка перехода к настройкам
        self.ui.btnHelp.clicked.connect(self.openDocumentation)
        self.ui.btnOpenRes.clicked.connect(self.openExplorer)
        self.ui.btnClearFiles.clicked.connect(self.clearFiles)

        self.doc_checked_count = 0

        self.thread = None

    def closeEvent(self, event):
        self.dlg = QDialogClass(self)
        if self.dlg.exec_() == QDialog.Accepted:
            event.accept()
        else:
            event.ignore()

    def saveSettings(self, geometry):
        self.settings.setValue("geometry", geometry)

        self.saveSettingsDefaultPage()
        self.saveSettingsDefaultHeadings()
        self.saveSettingsDefaultMainText()
        self.saveSettingsDefaultList()
        self.saveSettingsDefaultTable()
        self.saveSettingsDefaultPicture()

    def loadSettings(self):
        if self.settings.value("geometry"):
            self.setGeometry(self.settings.value("geometry"))
        self.SettingsWindow.setUserSettings()

    def saveSettingsDefaultPage(self):
        self.settings.setValue("LEFieldsBottom", self.SettingsWindow.ui.LEFieldsBottom.text())
        self.settings.setValue("LEFieldsTop", self.SettingsWindow.ui.LEFieldsTop.text())
        self.settings.setValue("LEFieldsLeft", self.SettingsWindow.ui.LEFieldsLeft.text())
        self.settings.setValue("LEFieldsRight", self.SettingsWindow.ui.LEFieldsRight.text())

        self.settings.setValue("CBFontName", self.SettingsWindow.ui.CBFontName.currentIndex())

        self.settings.setValue("PortraitOrientation", self.SettingsWindow.ui.PortraitOrientation.isChecked())
        self.settings.setValue("LandscapeOrientation", self.SettingsWindow.ui.LandscapeOrientation.isChecked())

    def saveSettingsDefaultHeadings(self):
        self.settings.setValue("LEFirstLvlSpacingAfter", self.SettingsWindow.ui.LEFirstLvlSpacingAfter.text())
        self.settings.setValue("LEFirstLvlSize", self.SettingsWindow.ui.LEFirstLvlSize.text())
        self.settings.setValue("LEFirstLvlSpacingBefore", self.SettingsWindow.ui.LEFirstLvlSpacingBefore.text())
        self.settings.setValue("LESecondLvlSpacingBefore", self.SettingsWindow.ui.LESecondLvlSpacingBefore.text())
        self.settings.setValue("LESecondLvlSize", self.SettingsWindow.ui.LESecondLvlSize.text())
        self.settings.setValue("LESecondLvlSpacingAfter", self.SettingsWindow.ui.LESecondLvlSpacingAfter.text())
        self.settings.setValue("LEThirdLvlSpacingBefore", self.SettingsWindow.ui.LEThirdLvlSpacingBefore.text())
        self.settings.setValue("LEThirdLvlSize", self.SettingsWindow.ui.LEThirdLvlSize.text())
        self.settings.setValue("LEThirdLvlSpacingAfter", self.SettingsWindow.ui.LEThirdLvlSpacingAfter.text())

        self.settings.setValue("CBLVL1CheckNumeration", self.SettingsWindow.ui.CBLVL1CheckNumeration.isChecked())
        self.settings.setValue("CBLVL1NotSpacing", self.SettingsWindow.ui.CBLVL1NotSpacing.isChecked())
        self.settings.setValue("CBLVL1NewPage", self.SettingsWindow.ui.CBLVL1NewPage.isChecked())

        self.settings.setValue("CBLVL2CheckNumeration", self.SettingsWindow.ui.CBLVL2CheckNumeration.isChecked())
        self.settings.setValue("CBLVL2NotSpacing", self.SettingsWindow.ui.CBLVL2NotSpacing.isChecked())
        self.settings.setValue("CBLVL2NewPage", self.SettingsWindow.ui.CBLVL2NewPage.isChecked())

        self.settings.setValue("CBLVL3CheckNumeration", self.SettingsWindow.ui.CBLVL3CheckNumeration.isChecked())
        self.settings.setValue("CBLVL3NotSpacing", self.SettingsWindow.ui.CBLVL3NotSpacing.isChecked())
        self.settings.setValue("CBLVL3NewPage", self.SettingsWindow.ui.CBLVL3NewPage.isChecked())

        self.settings.setValue("CBLVL1Bold", self.SettingsWindow.ui.CBLVL1Bold.isChecked())
        self.settings.setValue("CBLVL1Italic", self.SettingsWindow.ui.CBLVL1Italic.isChecked())
        self.settings.setValue("CBLVL1Underline", self.SettingsWindow.ui.CBLVL1Underline.isChecked())

        self.settings.setValue("CBLVL2Bold", self.SettingsWindow.ui.CBLVL2Bold.isChecked())
        self.settings.setValue("CBLVL2Italic", self.SettingsWindow.ui.CBLVL2Italic.isChecked())
        self.settings.setValue("CBLVL2Underline", self.SettingsWindow.ui.CBLVL2Underline.isChecked())

        self.settings.setValue("CBLVL3Bold", self.SettingsWindow.ui.CBLVL3Bold.isChecked())
        self.settings.setValue("CBLVL3Italic", self.SettingsWindow.ui.CBLVL3Italic.isChecked())
        self.settings.setValue("CBLVL3Underline", self.SettingsWindow.ui.CBLVL3Underline.isChecked())

        self.settings.setValue("RBLVL1TextLeft", self.SettingsWindow.ui.RBLVL1TextLeft.isChecked())
        self.settings.setValue("RBLVL1TextMiddle", self.SettingsWindow.ui.RBLVL1TextMiddle.isChecked())
        self.settings.setValue("RBLVL1TextRight", self.SettingsWindow.ui.RBLVL1TextRight.isChecked())
        self.settings.setValue("RBLVL1TextWidth", self.SettingsWindow.ui.RBLVL1TextWidth.isChecked())

        self.settings.setValue("RBLVL2TextLeft", self.SettingsWindow.ui.RBLVL2TextLeft.isChecked())
        self.settings.setValue("RBLVL2TextMiddle", self.SettingsWindow.ui.RBLVL2TextMiddle.isChecked())
        self.settings.setValue("RBLVL2TextRight", self.SettingsWindow.ui.RBLVL2TextRight.isChecked())
        self.settings.setValue("RBLVL2TextWidth", self.SettingsWindow.ui.RBLVL2TextWidth.isChecked())

        self.settings.setValue("RBLVL3TextLeft", self.SettingsWindow.ui.RBLVL3TextLeft.isChecked())
        self.settings.setValue("RBLVL3TextMiddle", self.SettingsWindow.ui.RBLVL3TextMiddle.isChecked())
        self.settings.setValue("RBLVL3TextRight", self.SettingsWindow.ui.RBLVL3TextRight.isChecked())
        self.settings.setValue("RBLVL3TextWidth", self.SettingsWindow.ui.RBLVL3TextWidth.isChecked())

    def saveSettingsDefaultMainText(self):
        self.settings.setValue("LEMainTextSpacingBefore", self.SettingsWindow.ui.LEMainTextSpacingBefore.text())
        self.settings.setValue("LEMainTextSpacingAfter", self.SettingsWindow.ui.LEMainTextSpacingAfter.text())
        self.settings.setValue("LEMainTextSize", self.SettingsWindow.ui.LEMainTextSize.text())
        self.settings.setValue("LEMainTextSpacingBetween", self.SettingsWindow.ui.LEMainTextSpacingBetween.text())
        self.settings.setValue("LEMainTextSpacingParagraph", self.SettingsWindow.ui.LEMainTextSpacingParagraph.text())

        self.settings.setValue("CBMainTextBold", self.SettingsWindow.ui.CBMainTextBold.isChecked())
        self.settings.setValue("CBMainTextItalic", self.SettingsWindow.ui.CBMainTextItalic.isChecked())
        self.settings.setValue("CBMainTextUnderline", self.SettingsWindow.ui.CBMainTextUnderline.isChecked())

        self.settings.setValue("RBMainTextLeft", self.SettingsWindow.ui.RBMainTextLeft.isChecked())
        self.settings.setValue("RBMainTextMiddle", self.SettingsWindow.ui.RBMainTextMiddle.isChecked())
        self.settings.setValue("RBMainTextRight", self.SettingsWindow.ui.RBMainTextRight.isChecked())
        self.settings.setValue("RBMainTextWidth", self.SettingsWindow.ui.RBMainTextWidth.isChecked())

    def saveSettingsDefaultList(self):
        self.settings.setValue("LEListMarginLeft", self.SettingsWindow.ui.LEListMarginLeft.text())
        self.settings.setValue("LEListMarginModify", self.SettingsWindow.ui.LEListMarginModify.text())
        self.settings.setValue("LEListLedge", self.SettingsWindow.ui.LEListLedge.text())
        self.settings.setValue("CBListReminder", self.SettingsWindow.ui.CBListReminder.isChecked())

    def saveSettingsDefaultTable(self):
        self.settings.setValue("LETableFontSize", self.SettingsWindow.ui.LETableFontSize.text())
        self.settings.setValue("CBTableParagraphBeforeTable", self.SettingsWindow.ui.CBTableParagraphBeforeTable.isChecked())
        self.settings.setValue("CBTableFormatParagraph", self.SettingsWindow.ui.CBTableFormatParagraph.currentIndex())

        self.settings.setValue("LETableSpacingBefore", self.SettingsWindow.ui.LETableSpacingBefore.text())
        self.settings.setValue("LETableSpacingAfter", self.SettingsWindow.ui.LETableSpacingAfter.text())
        self.settings.setValue("LETabletSpacingBetween", self.SettingsWindow.ui.LETabletSpacingBetween.text())
        self.settings.setValue("LETableSpacingParagraph", self.SettingsWindow.ui.LETableSpacingParagraph.text())
        self.settings.setValue("LETableParagraphSpacingAfter", self.SettingsWindow.ui.LETableParagraphSpacingAfter.text())

        self.settings.setValue("CBTableHeadingTop", self.SettingsWindow.ui.CBTableHeadingTop.isChecked())
        self.settings.setValue("CBTableHeadingLeft", self.SettingsWindow.ui.CBTableHeadingLeft.isChecked())

        self.settings.setValue("CBTableBold", self.SettingsWindow.ui.CBTableBold.isChecked())
        self.settings.setValue("CBTableItalic", self.SettingsWindow.ui.CBTableItalic.isChecked())
        self.settings.setValue("CBTableUnderline", self.SettingsWindow.ui.CBTableUnderline.isChecked())

        self.settings.setValue("RBTableHeadingTextLeft", self.SettingsWindow.ui.RBTableHeadingTextLeft.isChecked())
        self.settings.setValue("RBTableHeadingTextMiddle", self.SettingsWindow.ui.RBTableHeadingTextMiddle.isChecked())
        self.settings.setValue("RBTableHeadingTextRight", self.SettingsWindow.ui.RBLVL2TextRight.isChecked())

        self.settings.setValue("RBTableTextLeft", self.SettingsWindow.ui.RBTableTextLeft.isChecked())
        self.settings.setValue("RBTableTextMiddle", self.SettingsWindow.ui.RBTableTextMiddle.isChecked())
        self.settings.setValue("RBTableTextRight", self.SettingsWindow.ui.RBTableTextRight.isChecked())

        self.settings.setValue("RBTableTextTop", self.SettingsWindow.ui.RBTableTextTop.isChecked())
        self.settings.setValue("RBTableTextMiddle_2", self.SettingsWindow.ui.RBTableTextMiddle_2.isChecked())
        self.settings.setValue("RBTableTextBottom", self.SettingsWindow.ui.RBTableTextBottom.isChecked())

    def saveSettingsDefaultPicture(self):
        self.settings.setValue("LEPictureSpacingBefore", self.SettingsWindow.ui.LEPictureSpacingBefore.text())
        self.settings.setValue("LEPictureSpacingAfter", self.SettingsWindow.ui.LEPictureSpacingAfter.text())
        self.settings.setValue("LEPicturetSpacingParagraph", self.SettingsWindow.ui.LEPicturetSpacingParagraph.text())
        self.settings.setValue("LEPictureSpacingBetween", self.SettingsWindow.ui.LEPictureSpacingBetween.text())

        self.settings.setValue("CBPictureNotSpacing", self.SettingsWindow.ui.CBPictureNotSpacing.isChecked())

        self.settings.setValue("RBPictureLeft", self.SettingsWindow.ui.RBPictureLeft.isChecked())
        self.settings.setValue("RBPictureMiddle", self.SettingsWindow.ui.RBPictureMiddle.isChecked())
        self.settings.setValue("RBPictureRight", self.SettingsWindow.ui.RBPictureRight.isChecked())

        self.settings.setValue("CBPictureTitle", self.SettingsWindow.ui.CBPictureTitle.isChecked())
        self.settings.setValue("CBPictureTitleFormat", self.SettingsWindow.ui.CBPictureTitleFormat.currentIndex())

        self.settings.setValue("LEPictureFontSize", self.SettingsWindow.ui.LEPictureFontSize.text())
        self.settings.setValue("LEPictureTitleSpacingBefore", self.SettingsWindow.ui.LEPictureTitleSpacingBefore.text())
        self.settings.setValue("LEPictureTitleSpacingAfter", self.SettingsWindow.ui.LEPictureTitleSpacingAfter.text())
        self.settings.setValue("LEPictureTitleSpacingBetween", self.SettingsWindow.ui.LEPictureTitleSpacingBetween.text())
        self.settings.setValue("LEPictureTitleSpacingFirstLine", self.SettingsWindow.ui.LEPictureTitleSpacingFirstLine.text())

        self.settings.setValue("RBPictureTitleLeft", self.SettingsWindow.ui.RBPictureTitleLeft.isChecked())
        self.settings.setValue("RBPictureTitleMiddle", self.SettingsWindow.ui.RBPictureTitleMiddle.isChecked())
        self.settings.setValue("RBPictureTitleRight", self.SettingsWindow.ui.RBPictureTitleRight.isChecked())

        self.settings.setValue("CBPictureTitleUnderline", self.SettingsWindow.ui.CBPictureTitleUnderline.isChecked())
        self.settings.setValue("CBPictureTitleItalic", self.SettingsWindow.ui.CBPictureTitleItalic.isChecked())
        self.settings.setValue("CBPictureTitleBold", self.SettingsWindow.ui.CBPictureTitleBold.isChecked())

    def openDocumentation(self):
        if self.proc is not None:
            self.proc.kill()
        # 123
        menu_path = get_internal_data('Help Menu.chm')
        if self.isActiveWindow():
            self.proc = subprocess.Popen("hh.exe -mapid" + "100" + f" {menu_path}")
        else:
            self.proc = subprocess.Popen("hh.exe -mapid" + "20" + str(self.SettingsWindow.ui.tabWidget.currentIndex() + 1) + f" {menu_path}")

    def openExplorer(self):
        if self.explorer is not None:
            self.explorer.kill()
        if os.path.isdir("_internal"):
            self.explorer = subprocess.Popen(f'explorer "{self.file_path}\\..\\Results"')
        else:
            self.explorer = subprocess.Popen(f'explorer "{self.file_path}\\Results"')

    def chooseFile(self):  # Выбор файла
        self.ui.LEResLine.hide()
        self.ui.btnOpenRes.hide()
        dialog = QFileDialog.getOpenFileNames(self, "Выбор файла", "", "*.docx")

        if dialog[0]:
            for filename in dialog[0]:
                file, extension = os.path.splitext(filename)
                if extension == ".docx" and filename not in self.fileNames:
                    self.fileFlag = True
                    self.fileNames.append(filename)
                    self.ui.LENameFile.setText(f"Выбрано файлов: {len(self.fileNames)} ")
                    self.ui.LEErrorLine.hide()

    def clearFiles(self):
        self.ui.LENameFile.setText(f"Файлы не выбраны")
        self.fileFlag = False
        self.fileNames = []

    def update_progress_bar(self):
        self.doc_checked_count += 1
        self.ui.LEResLine.setText(f"Проверено файлов: {self.doc_checked_count}")

    def on_checking_finished(self):
        self.ui.btnChooseFile.setEnabled(True)
        self.ui.btnClearFiles.setEnabled(True)
        self.ui.btnRun.setEnabled(True)
        self.ui.btnOpenRes.setEnabled(True)
        self.ui.btnSettings.setEnabled(True)

        self.doc_checked_count = 0
        self.ui.btnOpenRes.show()
        self.clearFiles()
        self.thread.quit()

    def runCheckingCorrectness(self):  # Запуск проверки корректности
        if self.fileFlag:
            self.ui.btnChooseFile.setEnabled(False)
            self.ui.btnClearFiles.setEnabled(False)
            self.ui.btnRun.setEnabled(False)
            self.ui.btnOpenRes.setEnabled(False)
            self.ui.btnSettings.setEnabled(False)
            self.ui.LENameFile.clearFocus()

            self.doc_checked_count = 0
            self.getSettings()
            self.getOptionalSettings()
            parser = DocumentParser()
            parser.set_settings(self.text_checklist, self.heading1_checklist, self.heading2_checklist,
                                self.heading3_checklist, self.table_title_checklist, self.table_heading_checklist,
                                self.table_checklist, self.list_checklist, self.page_checklist, self.picture_checklist,
                                self.title_picture_checklist)
            parser.set_enable_optional_settings(self.enable_optional_settings)

            self.ui.LEResLine.setText(f"Проверено файлов: 0")
            self.ui.LEResLine.show()

            self.thread = Worker(self.fileNames, parser)
            self.thread.doc_progressed.connect(self.update_progress_bar)
            # self.thread.started.connect(self.thread.run)

            self.thread.start()
            self.thread.finished.connect(self.on_checking_finished)
        else:
            self.ui.LEErrorLine.show()  # Вывод сообщения об отсутствии выбранных файлов
            self.ui.LEResLine.hide()
            self.ui.btnOpenRes.hide()

    def getOptionalSettings(self):
        self.enable_optional_settings = {
            "table_headings_top": self.SettingsWindow.ui.CBTableHeadingTop.isChecked(),  # Необходимость заголовков сверху
            "table_headings_left": self.SettingsWindow.ui.CBTableHeadingLeft.isChecked(),  # Необходимость заголовков слева
            "table_title": self.SettingsWindow.ui.CBTableParagraphBeforeTable.isChecked(),  # Параграф перед таблицей
            "paragraph_after_table": True,
            "enable_pic_title": self.SettingsWindow.ui.CBPictureTitle.isChecked(),
            "list_reminder": self.SettingsWindow.ui.CBListReminder.isChecked()
        }

    def getSettings(self):
        self.getMainTextSettings()
        self.getHeading1Settings()
        self.getHeading2Settings()
        self.getHeading3Settings()
        self.getPageSettings()
        self.getListSettings()
        self.getPictureSettings()
        self.getTableSettings()

    def getMainTextSettings(self):
        alignment = 0
        if self.SettingsWindow.ui.RBMainTextLeft.isChecked():
            alignment = 0
        elif self.SettingsWindow.ui.RBMainTextRight.isChecked():
            alignment = 2
        elif self.SettingsWindow.ui.RBMainTextMiddle.isChecked():
            alignment = 1
        elif self.SettingsWindow.ui.RBMainTextWidth.isChecked():
            alignment = 3

        self.text_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),
            "font_size": float(self.SettingsWindow.ui.LEMainTextSize.text()),
            "font_bald": self.SettingsWindow.ui.CBMainTextBold.isChecked(),
            "font_italic": self.SettingsWindow.ui.CBMainTextItalic.isChecked(),
            "font_underline": self.SettingsWindow.ui.CBMainTextUnderline.isChecked(),
            "font_color": False,  # нет такого в UI
            "font_back_color": False,  # нет такого в UI
            "alignment": alignment,
            "keep_with_next": False,  # нет такого в UI, false по умолчанию для main текста
            "page_break_before": False,  # нет такого в UI, false по умолчанию для main текста
            "space_before": float(self.SettingsWindow.ui.LEMainTextSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LEMainTextSpacingAfter.text()),
            "left_indent": 0,  # нет такого в UI
            "right_indent": 0,  # нет такого в UI
            "first_line_indent": float(self.SettingsWindow.ui.LEMainTextSpacingParagraph.text()),
            "line_spacing": float(self.SettingsWindow.ui.LEMainTextSpacingBetween.text()),
        }

    def getHeading1Settings(self):
        alignment = 1
        if self.SettingsWindow.ui.RBLVL1TextLeft.isChecked():
            alignment = 0
        elif self.SettingsWindow.ui.RBLVL1TextRight.isChecked():
            alignment = 2
        elif self.SettingsWindow.ui.RBLVL1TextMiddle.isChecked():
            alignment = 1
        elif self.SettingsWindow.ui.RBLVL1TextLeft.isChecked():
            alignment = 3

        self.heading1_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),
            "font_size": float(self.SettingsWindow.ui.LEFirstLvlSize.text()),
            "font_bald": self.SettingsWindow.ui.CBLVL1Bold.isChecked(),
            "font_italic": self.SettingsWindow.ui.CBLVL1Italic.isChecked(),
            "font_underline": self.SettingsWindow.ui.CBLVL1Underline.isChecked(),
            "font_color": False,  # нет такого в UI
            "font_back_color": False,  # нет такого в UI
            "alignment": alignment,
            "keep_with_next": self.SettingsWindow.ui.CBLVL1NotSpacing.isChecked(),
            "page_break_before": self.SettingsWindow.ui.CBLVL1NewPage.isChecked(),  # нет такого в UI, false по умолчанию для main текста
            "space_before": float(self.SettingsWindow.ui.LEFirstLvlSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LEFirstLvlSpacingAfter.text()),
            "left_indent": 0,  # нет такого в UI
            "right_indent": 0,  # нет такого в UI
            "first_line_indent": 0,  # нет такого в UI
            "line_spacing": 1.0,  # нет такого в UI
            "is_list":  self.SettingsWindow.ui.CBLVL1CheckNumeration.isChecked()
        }

    def getHeading2Settings(self):
        alignment = 1
        if self.SettingsWindow.ui.RBLVL2TextLeft.isChecked():
            alignment = 0
        elif self.SettingsWindow.ui.RBLVL2TextRight.isChecked():
            alignment = 2
        elif self.SettingsWindow.ui.RBLVL2TextMiddle.isChecked():
            alignment = 1
        elif self.SettingsWindow.ui.RBLVL2TextLeft.isChecked():
            alignment = 3

        self.heading2_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),
            "font_size": float(self.SettingsWindow.ui.LESecondLvlSize.text()),
            "font_bald": self.SettingsWindow.ui.CBLVL2Bold.isChecked(),
            "font_italic": self.SettingsWindow.ui.CBLVL2Italic.isChecked(),
            "font_underline": self.SettingsWindow.ui.CBLVL2Underline.isChecked(),
            "font_color": False,  # нет такого в UI
            "font_back_color": False,  # нет такого в UI
            "alignment": alignment,
            "keep_with_next": self.SettingsWindow.ui.CBLVL2NotSpacing.isChecked(),
            "page_break_before": self.SettingsWindow.ui.CBLVL2NewPage.isChecked(),  # нет такого в UI, false по умолчанию для main текста
            "space_before": float(self.SettingsWindow.ui.LESecondLvlSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LESecondLvlSpacingAfter.text()),
            "left_indent": 0,  # нет такого в UI
            "right_indent": 0,  # нет такого в UI
            "first_line_indent": 0,  # нет такого в UI
            "line_spacing": 1.0,  # нет такого в UI
            "is_list": self.SettingsWindow.ui.CBLVL2CheckNumeration.isChecked()
        }

    def getHeading3Settings(self):
        alignment = 1
        if self.SettingsWindow.ui.RBLVL3TextLeft.isChecked():
            alignment = 0
        elif self.SettingsWindow.ui.RBLVL3TextRight.isChecked():
            alignment = 2
        elif self.SettingsWindow.ui.RBLVL3TextMiddle.isChecked():
            alignment = 1
        elif self.SettingsWindow.ui.RBLVL3TextLeft.isChecked():
            alignment = 3

        self.heading3_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),
            "font_size": float(self.SettingsWindow.ui.LEThirdLvlSize.text()),
            "font_bald": self.SettingsWindow.ui.CBLVL3Bold.isChecked(),
            "font_italic": self.SettingsWindow.ui.CBLVL3Italic.isChecked(),
            "font_underline": self.SettingsWindow.ui.CBLVL3Underline.isChecked(),
            "font_color": False,  # нет такого в UI
            "font_back_color": False,  # нет такого в UI
            "alignment": alignment,
            "keep_with_next": self.SettingsWindow.ui.CBLVL3NotSpacing.isChecked(),
            "page_break_before": self.SettingsWindow.ui.CBLVL3NewPage.isChecked(),  # нет такого в UI, false по умолчанию для main текста
            "space_before": float(self.SettingsWindow.ui.LEThirdLvlSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LEThirdLvlSpacingAfter.text()),
            "left_indent": 0,  # нет такого в UI
            "right_indent": 0,  # нет такого в UI
            "first_line_indent": 0,  # нет такого в UI
            "line_spacing": 1.0,  # нет такого в UI
            "is_list": self.SettingsWindow.ui.CBLVL3CheckNumeration.isChecked()
        }

    def getPageSettings(self):

        orientation = 0
        if self.SettingsWindow.ui.LandscapeOrientation.isChecked():
            orientation = 1

        self.page_checklist = {
            "top_margin": float(self.SettingsWindow.ui.LEFieldsTop.text()),  # Поля страницы (верхнее)
            "bottom_margin": float(self.SettingsWindow.ui.LEFieldsBottom.text()),  # Поля страницы (нижнее)
            "left_margin": float(self.SettingsWindow.ui.LEFieldsLeft.text()),  # Поля страницы (левое)
            "right_margin": float(self.SettingsWindow.ui.LEFieldsRight.text()),  # Поля страницы (правое)
            "orientation": orientation
        }

    def getListSettings(self):
        self.list_checklist = {
            "left_indent_base": float(self.SettingsWindow.ui.LEListMarginLeft.text()),
            "left_indent_mod": float(self.SettingsWindow.ui.LEListMarginModify.text()),
            "first_line_indent": -float(self.SettingsWindow.ui.LEListLedge.text()),
        }

    def getTableSettings(self):
        alignment = 0
        if self.SettingsWindow.ui.RBTableTextLeft.isChecked():
            alignment = 0
        elif self.SettingsWindow.ui.RBTableTextRight.isChecked():
            alignment = 2
        elif self.SettingsWindow.ui.RBLVL3TextMiddle.isChecked():
            alignment = 1

        vertical_alignment = 1
        if self.SettingsWindow.ui.RBTableTextTop.isChecked():
            vertical_alignment = 0
        elif self.SettingsWindow.ui.RBTableTextMiddle_2.isChecked():
            vertical_alignment = 1
        elif self.SettingsWindow.ui.RBTableTextBottom.isChecked():
            vertical_alignment = 2

        heading_alignment = 1
        if self.SettingsWindow.ui.RBTableHeadingTextLeft.isChecked():
            heading_alignment = 0
        elif self.SettingsWindow.ui.RBTableHeadingTextRight.isChecked():
            heading_alignment = 2
        elif self.SettingsWindow.ui.RBTableHeadingTextMiddle.isChecked():
            heading_alignment = 1

        self.table_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),  # Шрифт
            "font_size": float(self.SettingsWindow.ui.LETableFontSize.text()),  # Размер шрифта в таблице
            "space_before": float(self.SettingsWindow.ui.LETableSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LETableSpacingAfter.text()),
            "left_indent": 0,  # нет такого в UI
            "right_indent": 0,  # нет такого в UI
            "first_line_indent": float(self.SettingsWindow.ui.LETableSpacingParagraph.text()),
            "line_spacing": float(self.SettingsWindow.ui.LETabletSpacingBetween.text()),
            "spacing_under_paragraph_after_table": float(self.SettingsWindow.ui.LETableParagraphSpacingAfter.text()),  # Интервал абзаца после таблицы
            "alignment": alignment,
            "vertical_alignment": vertical_alignment
        }

        self.table_title_checklist = {
            "format_regex": self.SettingsWindow.ui.CBTableFormatParagraph.currentText(), # Формат подписи под таблицей (Таблица <N> - <Название>)
        }

        self.table_heading_checklist = {
            "font_bald": self.SettingsWindow.ui.CBTableBold.isChecked(),  # Жирный шрифт заголовков
            "font_italic": self.SettingsWindow.ui.CBTableItalic.isChecked(),  # Курсив заголовков
            "font_underline": self.SettingsWindow.ui.CBTableUnderline.isChecked(),  # Подчеркивание заголовков
            "alignment": heading_alignment,
            "vertical_alignment": vertical_alignment
            }

    def getPictureSettings(self):
        title_alignment = 0
        if self.SettingsWindow.ui.RBPictureTitleLeft.isChecked():
            title_alignment = 0
        elif self.SettingsWindow.ui.RBPictureTitleRight.isChecked():
            title_alignment = 2
        elif self.SettingsWindow.ui.RBPictureTitleMiddle.isChecked():
            title_alignment = 1

        picture_alignment = 1
        if self.SettingsWindow.ui.RBPictureLeft.isChecked():
            picture_alignment = 0
        elif self.SettingsWindow.ui.RBPictureRight.isChecked():
            picture_alignment = 2
        elif self.SettingsWindow.ui.RBPictureMiddle.isChecked():
            picture_alignment = 1

        self.picture_checklist = {
            "alignment": picture_alignment,  # Выравнивание картинки
            "keep_with_next": self.SettingsWindow.ui.CBPictureNotSpacing.isChecked(),  # Не отрывать рисунок от подписи
            "space_before": float(self.SettingsWindow.ui.LEPictureSpacingBefore.text()),
            "space_after": float(self.SettingsWindow.ui.LEPictureSpacingAfter.text()),
            "first_line_indent": float(self.SettingsWindow.ui.LEPicturetSpacingParagraph.text()),
            "line_spacing": float(self.SettingsWindow.ui.LEPictureSpacingBetween.text()),
        }

        self.title_picture_checklist = {
            "font_name": self.SettingsWindow.ui.CBFontName.currentText(),  # Шрифт подписи под рисунком
            "font_size": float(self.SettingsWindow.ui.LEPictureFontSize.text()),  # Размер подписи под рисунком
            "font_bald": self.SettingsWindow.ui.CBPictureTitleBold.isChecked(),  # Выделение жирным шрифтом
            "font_italic": self.SettingsWindow.ui.CBPictureTitleItalic.isChecked(),  # Выделение курсовом
            "font_underline": self.SettingsWindow.ui.CBPictureTitleUnderline.isChecked(),  # Выделение подчеркиванием
            "space_before": float(self.SettingsWindow.ui.LEPictureTitleSpacingBefore.text()),  # интервал перед подписью
            "space_after": float(self.SettingsWindow.ui.LEPictureTitleSpacingAfter.text()),  # интервал после подписи
            "first_line_indent": float(self.SettingsWindow.ui.LEPictureTitleSpacingFirstLine.text()),  # Абзацный отступ
            "line_spacing": float(self.SettingsWindow.ui.LEPictureTitleSpacingBetween.text()),  # Междустрочный интервал
            "alignment": title_alignment,  # Выравнивание подписи
            "format_regex": self.SettingsWindow.ui.CBPictureTitleFormat.currentText()  # Формат подписи (Рисунок <N> - <Название>.)
        }

    def setupSettings(self):  # Запуск окна настроек
        self.hide()
        self.SettingsWindow.setGeometry(self.geometry())
        self.SettingsWindow.show()


def except_hook(cls, exception, traceback):  # Блок для получения сообщений об ошибках
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':  # Запуск программы
    multiprocessing.freeze_support()
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    sys.excepthook = except_hook

    sys.exit(app.exec())
