import sys

from PyQt5 import (uic, QtCore, QtWidgets)
from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QIntValidator, QDoubleValidator, QRegExpValidator
from PyQt5.QtWidgets import QWidget, QApplication, QDialog

from dialog import QDialogClass

from ui_settingswindow import Ui_Settings


class SettingsWindow(QWidget):
    def __init__(self, *args):
        super(SettingsWindow, self).__init__()
        self.ui = Ui_Settings()
        self.ui.setupUi(self)
        # self.ui = uic.loadUi('SettingsUI.ui', self)  # Открытие файла ui
        self.main = args[0]
        self.ui.btnGoBack.clicked.connect(self.goBack)  # Кнопка перехода обратно в основное меню
        self.ui.btnDefaultSettings.clicked.connect(self.setDefaultSettingsByScreen)  # Кнопка установки базовых настроек
        self.ui.btnHelp.clicked.connect(self.main.openDocumentation)
        self.ui.tabWidget.setCurrentIndex(0)
        self.dlg = None
        self.setGeometry(self.main.geometry())

        self.setWindowTitle("HSEReportChecker")

        self.setValidators()

    def goBack(self):  # Вернуться назад в основное окно
        self.hide()
        self.main.setGeometry(self.geometry())
        self.main.show()

    def closeEvent(self, event):
        self.dlg = QDialogClass(self, self.main)
        if self.dlg.exec_() == QDialog.Accepted:
            event.accept()
        else:
            event.ignore()

    def setUserSettings(self):  # Установка базовых настроек
        self.setUserSettingsPage()
        self.setUserSettingsHeadings()
        self.setUserSettingsMainText()
        self.setUserSettingsList()
        self.setUserSettingsTable()
        self.setUserSettingsPicture()

    def setUserSettingsPage(self):
        self.ui.LEFieldsBottom.setText(self.main.settings.value("LEFieldsBottom", "2"))
        self.ui.LEFieldsTop.setText(self.main.settings.value("LEFieldsTop", "2"))
        self.ui.LEFieldsLeft.setText(self.main.settings.value("LEFieldsLeft", "3"))
        self.ui.LEFieldsRight.setText(self.main.settings.value("LEFieldsRight", "1.5"))

        self.ui.CBFontName.setCurrentIndex(self.main.settings.value("CBFontName", 0))

        self.ui.PortraitOrientation.setChecked(self.main.settings.value("PortraitOrientation", True, type=bool))
        self.ui.LandscapeOrientation.setChecked(self.main.settings.value("LandscapeOrientation", False, type=bool))

    def setUserSettingsHeadings(self):
        self.ui.LEFirstLvlSpacingAfter.setText(self.main.settings.value("LEFirstLvlSpacingAfter", "12"))
        self.ui.LEFirstLvlSize.setText(self.main.settings.value("LEFirstLvlSize", "16"))
        self.ui.LEFirstLvlSpacingBefore.setText(self.main.settings.value("LEFirstLvlSpacingBefore", "0"))
        self.ui.LESecondLvlSpacingBefore.setText(self.main.settings.value("LESecondLvlSpacingBefore", "12"))
        self.ui.LESecondLvlSize.setText(self.main.settings.value("LESecondLvlSize", "14"))
        self.ui.LESecondLvlSpacingAfter.setText(self.main.settings.value("LESecondLvlSpacingAfter", "6"))
        self.ui.LEThirdLvlSpacingBefore.setText(self.main.settings.value("LEThirdLvlSpacingBefore", "8"))
        self.ui.LEThirdLvlSize.setText(self.main.settings.value("LEThirdLvlSize", "13"))
        self.ui.LEThirdLvlSpacingAfter.setText(self.main.settings.value("LEThirdLvlSpacingAfter", "4"))

        self.ui.CBLVL1CheckNumeration.setChecked(self.main.settings.value("CBLVL1CheckNumeration", True, type=bool))
        self.ui.CBLVL1NotSpacing.setChecked(self.main.settings.value("CBLVL1NotSpacing", True, type=bool))
        self.ui.CBLVL1NewPage.setChecked(self.main.settings.value("CBLVL1NewPage", True, type=bool))

        self.ui.CBLVL2CheckNumeration.setChecked(self.main.settings.value("CBLVL2CheckNumeration", True, type=bool))
        self.ui.CBLVL2NotSpacing.setChecked(self.main.settings.value("CBLVL2NotSpacing", True, type=bool))
        self.ui.CBLVL2NewPage.setChecked(self.main.settings.value("CBLVL2NewPage", False, type=bool))

        self.ui.CBLVL3CheckNumeration.setChecked(self.main.settings.value("CBLVL3CheckNumeration", True, type=bool))
        self.ui.CBLVL3NotSpacing.setChecked(self.main.settings.value("CBLVL3NotSpacing", True, type=bool))
        self.ui.CBLVL3NewPage.setChecked(self.main.settings.value("CBLVL3NewPage", False, type=bool))

        self.ui.CBLVL1Bold.setChecked(self.main.settings.value("CBLVL1Bold", True, type=bool))
        self.ui.CBLVL1Italic.setChecked(self.main.settings.value("CBLVL1Italic", False, type=bool))
        self.ui.CBLVL1Underline.setChecked(self.main.settings.value("CBLVL1Underline", False, type=bool))

        self.ui.CBLVL2Bold.setChecked(self.main.settings.value("CBLVL2Bold", True, type=bool))
        self.ui.CBLVL2Italic.setChecked(self.main.settings.value("CBLVL2Italic", False, type=bool))
        self.ui.CBLVL2Underline.setChecked(self.main.settings.value("CBLVL2Underline", False, type=bool))

        self.ui.CBLVL3Bold.setChecked(self.main.settings.value("CBLVL3Bold.", True, type=bool))
        self.ui.CBLVL3Italic.setChecked(self.main.settings.value("CBLVL3Italic", False, type=bool))
        self.ui.CBLVL3Underline.setChecked(self.main.settings.value("CBLVL3Underline", False, type=bool))

        self.ui.RBLVL1TextLeft.setChecked(self.main.settings.value("RBLVL1TextLeft", False, type=bool))
        self.ui.RBLVL1TextMiddle.setChecked(self.main.settings.value("RBLVL1TextMiddle", True, type=bool))
        self.ui.RBLVL1TextRight.setChecked(self.main.settings.value("RBLVL1TextRight", False, type=bool))
        self.ui.RBLVL1TextWidth.setChecked(self.main.settings.value("RBLVL1TextWidth", False, type=bool))

        self.ui.RBLVL2TextLeft.setChecked(self.main.settings.value("RBLVL2TextLeft", False, type=bool))
        self.ui.RBLVL2TextMiddle.setChecked(self.main.settings.value("RBLVL2TextMiddle", True, type=bool))
        self.ui.RBLVL2TextRight.setChecked(self.main.settings.value("RBLVL2TextRight", False, type=bool))
        self.ui.RBLVL2TextWidth.setChecked(self.main.settings.value("RBLVL2TextWidth", False, type=bool))

        self.ui.RBLVL3TextLeft.setChecked(self.main.settings.value("RBLVL3TextLeft", False, type=bool))
        self.ui.RBLVL3TextMiddle.setChecked(self.main.settings.value("RBLVL3TextMiddle", True, type=bool))
        self.ui.RBLVL3TextRight.setChecked(self.main.settings.value("RBLVL3TextRight", False, type=bool))
        self.ui.RBLVL3TextWidth.setChecked(self.main.settings.value("RBLVL3TextWidth", False, type=bool))

    def setUserSettingsMainText(self):
        self.ui.LEMainTextSpacingBefore.setText(self.main.settings.value("LEMainTextSpacingBefore", "0"))
        self.ui.LEMainTextSpacingAfter.setText(self.main.settings.value("LEMainTextSpacingAfter", "0"))
        self.ui.LEMainTextSize.setText(self.main.settings.value("LEMainTextSize", "13"))
        self.ui.LEMainTextSpacingBetween.setText(self.main.settings.value("LEMainTextSpacingBetween", "1.5"))
        self.ui.LEMainTextSpacingParagraph.setText(self.main.settings.value("LEMainTextSpacingParagraph", "1.25"))

        self.ui.CBMainTextBold.setChecked(self.main.settings.value("CBMainTextBold", False, type=bool))
        self.ui.CBMainTextItalic.setChecked(self.main.settings.value("CBMainTextItalic", False, type=bool))
        self.ui.CBMainTextUnderline.setChecked(self.main.settings.value("CBMainTextUnderline", False, type=bool))

        self.ui.RBMainTextLeft.setChecked(self.main.settings.value("RBMainTextLeft", False, type=bool))
        self.ui.RBMainTextMiddle.setChecked(self.main.settings.value("RBMainTextMiddle", False, type=bool))
        self.ui.RBMainTextRight.setChecked(self.main.settings.value("RBMainTextRight", False, type=bool))
        self.ui.RBMainTextWidth.setChecked(self.main.settings.value("RBMainTextWidth", True, type=bool))

    def setUserSettingsList(self):
        self.ui.LEListMarginLeft.setText(self.main.settings.value("LEListMarginLeft", "1.25"))
        self.ui.LEListMarginModify.setText(self.main.settings.value("LEListMarginModify", "0.75"))
        self.ui.LEListLedge.setText(self.main.settings.value("LEListLedge", "0.75"))
        self.ui.CBListReminder.setChecked(self.main.settings.value("CBListReminder", True, type=bool))

    def setUserSettingsTable(self):
        self.ui.LETableFontSize.setText(self.main.settings.value("LETableFontSize", "12"))
        self.ui.CBTableParagraphBeforeTable.setChecked(self.main.settings.value("CBTableParagraphBeforeTable", True, type=bool))
        self.ui.CBTableFormatParagraph.setCurrentIndex(self.main.settings.value("CBTableFormatParagraph", 0))

        self.ui.LETableSpacingBefore.setText(self.main.settings.value("LETableSpacingBefore", "13"))
        self.ui.LETableSpacingAfter.setText(self.main.settings.value("LETableSpacingAfter", "0"))
        self.ui.LETabletSpacingBetween.setText(self.main.settings.value("LETabletSpacingBetween", "1"))
        self.ui.LETableSpacingParagraph.setText(self.main.settings.value("LETableSpacingParagraph", "0"))
        self.ui.LETableParagraphSpacingAfter.setText(self.main.settings.value("LETableParagraphSpacingAfter", "13"))

        self.ui.CBTableHeadingTop.setChecked(self.main.settings.value("CBTableHeadingTop.", True, type=bool))
        self.ui.CBTableHeadingLeft.setChecked(self.main.settings.value("CBTableHeadingLeft", False, type=bool))

        self.ui.CBTableBold.setChecked(self.main.settings.value("CBTableBold", True, type=bool))
        self.ui.CBTableItalic.setChecked(self.main.settings.value("CBTableItalic", False, type=bool))
        self.ui.CBTableUnderline.setChecked(self.main.settings.value("CBTableUnderline", False, type=bool))

        self.ui.RBTableHeadingTextLeft.setChecked(self.main.settings.value("RBTableHeadingTextLeft", False, type=bool))
        self.ui.RBTableHeadingTextMiddle.setChecked(self.main.settings.value("RBTableHeadingTextMiddle", True, type=bool))
        self.ui.RBTableHeadingTextRight.setChecked(self.main.settings.value("RBTableHeadingTextRight", False, type=bool))

        self.ui.RBTableTextLeft.setChecked(self.main.settings.value("RBTableTextLeft", True, type=bool))
        self.ui.RBTableTextMiddle.setChecked(self.main.settings.value("RBTableTextMiddle", False, type=bool))
        self.ui.RBTableTextRight.setChecked(self.main.settings.value("RBTableTextRight", False, type=bool))

        self.ui.RBTableTextTop.setChecked(self.main.settings.value("RBTableTextTop", False, type=bool))
        self.ui.RBTableTextMiddle_2.setChecked(self.main.settings.value("RBTableTextMiddle_2", True, type=bool))
        self.ui.RBTableTextBottom.setChecked(self.main.settings.value("RBTableTextBottom", False, type=bool))

    def setUserSettingsPicture(self):
        self.ui.LEPictureSpacingBefore.setText(self.main.settings.value("LEPictureSpacingBefore", "6"))
        self.ui.LEPictureSpacingAfter.setText(self.main.settings.value("LEPictureSpacingAfter", "0"))
        self.ui.LEPicturetSpacingParagraph.setText(self.main.settings.value("LEPicturetSpacingParagraph", "0"))
        self.ui.LEPictureSpacingBetween.setText(self.main.settings.value("LEPictureSpacingBetween", "1"))

        self.ui.CBPictureNotSpacing.setChecked(self.main.settings.value("CBPictureNotSpacing", True, type=bool))

        self.ui.RBPictureLeft.setChecked(self.main.settings.value("RBPictureLeft", False, type=bool))
        self.ui.RBPictureMiddle.setChecked(self.main.settings.value("RBPictureMiddle", True, type=bool))
        self.ui.RBPictureRight.setChecked(self.main.settings.value("RBPictureRight", False, type=bool))

        self.ui.CBPictureTitle.setChecked(self.main.settings.value("CBPictureTitle", True, type=bool))

        self.ui.CBPictureTitleFormat.setCurrentIndex(self.main.settings.value("CBPictureTitleFormat", 0, type=int))

        self.ui.LEPictureFontSize.setText(self.main.settings.value("LEPictureFontSize", "1"))
        self.ui.LEPictureTitleSpacingBefore.setText(self.main.settings.value("LEPictureTitleSpacingBefore", "0"))
        self.ui.LEPictureTitleSpacingAfter.setText(self.main.settings.value("LEPictureTitleSpacingAfter", "6"))
        self.ui.LEPictureTitleSpacingBetween.setText(self.main.settings.value("LEPictureTitleSpacingBetween", "1"))
        self.ui.LEPictureTitleSpacingFirstLine.setText(self.main.settings.value("LEPictureTitleSpacingFirstLine", "0"))

        self.ui.RBPictureTitleLeft.setChecked(self.main.settings.value("RBPictureTitleLeft", False, type=bool))
        self.ui.RBPictureTitleMiddle.setChecked(self.main.settings.value("RBPictureTitleMiddle", True, type=bool))
        self.ui.RBPictureTitleRight.setChecked(self.main.settings.value("RBPictureTitleRight", False, type=bool))

        self.ui.CBPictureTitleUnderline.setChecked(self.main.settings.value("CBPictureTitleUnderline", False, type=bool))
        self.ui.CBPictureTitleItalic.setChecked(self.main.settings.value("CBPictureTitleItalic", True, type=bool))
        self.ui.CBPictureTitleBold.setChecked(self.main.settings.value("CBPictureTitleBold", True, type=bool))

    def setDefaultSettingsByScreen(self):
        lst = [self.setDefaultPage,
               self.setDefaultHeadings,
               self.setDefaultMainText,
               self.setDefaultList,
               self.setDefaultTable,
               self.setDefaultPicture]
        lst[self.ui.tabWidget.currentIndex()]()

    def setDefaultPage(self):
        self.ui.LEFieldsBottom.setText("2")
        self.ui.LEFieldsTop.setText("2")
        self.ui.LEFieldsLeft.setText("3")
        self.ui.LEFieldsRight.setText("1.5")

        self.ui.CBFontName.setCurrentIndex(0)
        self.ui.PortraitOrientation.setChecked(True)

    def setDefaultHeadings(self):
        self.ui.LEFirstLvlSpacingAfter.setText("12")
        self.ui.LEFirstLvlSize.setText("16")
        self.ui.LEFirstLvlSpacingBefore.setText("0")
        self.ui.LESecondLvlSpacingBefore.setText("12")
        self.ui.LESecondLvlSize.setText("14")
        self.ui.LESecondLvlSpacingAfter.setText("6")
        self.ui.LEThirdLvlSpacingBefore.setText("8")
        self.ui.LEThirdLvlSize.setText("13")
        self.ui.LEThirdLvlSpacingAfter.setText("4")

        self.ui.CBLVL1CheckNumeration.setChecked(True)
        self.ui.CBLVL2CheckNumeration.setChecked(True)
        self.ui.CBLVL3CheckNumeration.setChecked(True)

        self.ui.CBLVL1NotSpacing.setChecked(True)
        self.ui.CBLVL1NewPage.setChecked(True)
        self.ui.CBLVL2NotSpacing.setChecked(True)
        self.ui.CBLVL2NewPage.setChecked(False)
        self.ui.CBLVL3NewPage.setChecked(False)
        self.ui.CBLVL3NotSpacing.setChecked(True)
        self.ui.CBLVL1Bold.setChecked(True)
        self.ui.CBLVL1Italic.setChecked(False)
        self.ui.CBLVL1Underline.setChecked(False)
        self.ui.RBLVL1TextMiddle.setChecked(True)
        self.ui.CBLVL2Bold.setChecked(True)
        self.ui.CBLVL2Italic.setChecked(False)
        self.ui.CBLVL2Underline.setChecked(False)
        self.ui.RBLVL2TextMiddle.setChecked(True)
        self.ui.CBLVL3Bold.setChecked(True)
        self.ui.CBLVL3Italic.setChecked(False)
        self.ui.CBLVL3Underline.setChecked(False)
        self.ui.RBLVL3TextMiddle.setChecked(True)

    def setDefaultMainText(self):
        self.ui.LEMainTextSpacingBefore.setText("0")
        self.ui.LEMainTextSpacingAfter.setText("0")
        self.ui.LEMainTextSize.setText("13")
        self.ui.LEMainTextSpacingBetween.setText("1.5")
        self.ui.LEMainTextSpacingParagraph.setText("1.25")

        self.ui.CBMainTextBold.setChecked(False)
        self.ui.CBMainTextItalic.setChecked(False)
        self.ui.CBMainTextUnderline.setChecked(False)
        self.ui.RBMainTextWidth.setChecked(True)

    def setDefaultList(self):
        self.ui.LEListMarginLeft.setText("1.25")
        self.ui.LEListMarginModify.setText("0.75")
        self.ui.LEListLedge.setText("0.75")
        self.ui.CBListReminder.setChecked(True)

    def setDefaultTable(self):
        self.ui.LETableFontSize.setText("12")
        self.ui.CBTableParagraphBeforeTable.setChecked(True)
        self.ui.CBTableFormatParagraph.setCurrentIndex(0)

        self.ui.LETableSpacingBefore.setText("13")
        self.ui.LETableSpacingAfter.setText("0")
        self.ui.LETabletSpacingBetween.setText("1")
        self.ui.LETableSpacingParagraph.setText("0")
        self.ui.LETableParagraphSpacingAfter.setText("13")

        self.ui.CBTableHeadingTop.setChecked(True)
        self.ui.CBTableHeadingLeft.setChecked(False)

        self.ui.CBTableBold.setChecked(True)
        self.ui.CBTableItalic.setChecked(False)
        self.ui.CBTableUnderline.setChecked(False)
        self.ui.RBTableTextMiddle_2.setChecked(True)
        self.ui.RBTableTextLeft.setChecked(True)
        self.ui.RBTableHeadingTextMiddle.setChecked(True)

    def setDefaultPicture(self):
        self.ui.LEPictureSpacingBefore.setText("6")
        self.ui.LEPictureSpacingAfter.setText("0")
        self.ui.LEPicturetSpacingParagraph.setText("0")
        self.ui.LEPictureSpacingBetween.setText("1")

        self.ui.CBPictureNotSpacing.setChecked(True)
        self.ui.CBPictureTitle.setChecked(True)

        self.ui.CBPictureTitleFormat.setCurrentIndex(0)
        self.ui.LEPictureFontSize.setText("11")
        self.ui.LEPictureTitleSpacingBefore.setText("0")
        self.ui.LEPictureTitleSpacingAfter.setText("6")
        self.ui.LEPictureTitleSpacingBetween.setText("1")
        self.ui.LEPictureTitleSpacingFirstLine.setText("0")

        self.ui.RBPictureMiddle.setChecked(True)
        self.ui.RBPictureTitleMiddle.setChecked(True)
        self.ui.CBPictureTitleUnderline.setChecked(False)
        self.ui.CBPictureTitleItalic.setChecked(True)
        self.ui.CBPictureTitleBold.setChecked(True)

    def setValidators(self):
        regex = QRegExpValidator(QRegExp("^([1-9][0-9]?|0)(\\.)[0-9]{2}$"))
        self.ui.LEFieldsBottom.setValidator(regex)
        self.ui.LEFieldsTop.setValidator(regex)
        self.ui.LEFieldsLeft.setValidator(regex)
        self.ui.LEFieldsRight.setValidator(regex)

        self.ui.LEFirstLvlSpacingAfter.setValidator(regex)
        self.ui.LEFirstLvlSize.setValidator(regex)
        self.ui.LEFirstLvlSpacingBefore.setValidator(regex)
        self.ui.LESecondLvlSpacingBefore.setValidator(regex)
        self.ui.LESecondLvlSize.setValidator(regex)
        self.ui.LESecondLvlSpacingAfter.setValidator(regex)
        self.ui.LEThirdLvlSpacingBefore.setValidator(regex)
        self.ui.LEThirdLvlSize.setValidator(regex)
        self.ui.LEThirdLvlSpacingAfter.setValidator(regex)

        self.ui.LEMainTextSpacingBefore.setValidator(regex)
        self.ui.LEMainTextSpacingAfter.setValidator(regex)
        self.ui.LEMainTextSize.setValidator(regex)
        self.ui.LEMainTextSpacingBetween.setValidator(regex)
        self.ui.LEMainTextSpacingParagraph.setValidator(regex)

        self.ui.LEListMarginLeft.setValidator(regex)
        self.ui.LEListMarginModify.setValidator(regex)

        self.ui.LETableFontSize.setValidator(regex)

        self.ui.LETableSpacingBefore.setValidator(regex)
        self.ui.LETableSpacingAfter.setValidator(regex)
        self.ui.LETabletSpacingBetween.setValidator(regex)
        self.ui.LETableSpacingParagraph.setValidator(regex)
        self.ui.LETableParagraphSpacingAfter.setValidator(regex)

        self.ui.LEPictureSpacingBefore.setValidator(regex)
        self.ui.LEPictureSpacingAfter.setValidator(regex)
        self.ui.LEPicturetSpacingParagraph.setValidator(regex)
        self.ui.LEPictureSpacingBetween.setValidator(regex)
        self.ui.LEPictureFontSize.setValidator(regex)
        self.ui.LEPictureTitleSpacingBefore.setValidator(regex)
        self.ui.LEPictureTitleSpacingAfter.setValidator(regex)
        self.ui.LEPictureTitleSpacingBetween.setValidator(regex)
        self.ui.LEPictureTitleSpacingFirstLine.setValidator(regex)


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)
