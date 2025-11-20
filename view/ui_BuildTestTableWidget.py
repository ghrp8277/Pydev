# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'BuildTestTableWidget.ui'
##
## Created by: Qt User Interface Compiler version 6.10.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QHeaderView, QLabel,
    QPlainTextEdit, QSizePolicy, QTableWidget, QTableWidgetItem,
    QTextBrowser, QVBoxLayout, QWidget)

class Ui_BuildTestTableWidget(object):
    def setupUi(self, BuildTestTableWidget):
        if not BuildTestTableWidget.objectName():
            BuildTestTableWidget.setObjectName(u"BuildTestTableWidget")
        BuildTestTableWidget.resize(1371, 617)
        self.verticalLayout = QVBoxLayout(BuildTestTableWidget)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.DescriptionWidget = QWidget(BuildTestTableWidget)
        self.DescriptionWidget.setObjectName(u"DescriptionWidget")
        self.DescriptionWidget.setAutoFillBackground(True)
        self.verticalLayout_2 = QVBoxLayout(self.DescriptionWidget)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.label = QLabel(self.DescriptionWidget)
        self.label.setObjectName(u"label")

        self.verticalLayout_2.addWidget(self.label)

        self.plainTextEdit = QPlainTextEdit(self.DescriptionWidget)
        self.plainTextEdit.setObjectName(u"plainTextEdit")
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.plainTextEdit.sizePolicy().hasHeightForWidth())
        self.plainTextEdit.setSizePolicy(sizePolicy)
        self.plainTextEdit.setMaximumSize(QSize(16777215, 30))
        self.plainTextEdit.setAcceptDrops(True)
        self.plainTextEdit.setInputMethodHints(Qt.ImhNone)

        self.verticalLayout_2.addWidget(self.plainTextEdit)


        self.verticalLayout.addWidget(self.DescriptionWidget)

        self.tableWidget = QTableWidget(BuildTestTableWidget)
        if (self.tableWidget.columnCount() < 11):
            self.tableWidget.setColumnCount(11)
        __qtablewidgetitem = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, __qtablewidgetitem)
        __qtablewidgetitem1 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, __qtablewidgetitem1)
        __qtablewidgetitem2 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, __qtablewidgetitem2)
        __qtablewidgetitem3 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, __qtablewidgetitem3)
        __qtablewidgetitem4 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(4, __qtablewidgetitem4)
        __qtablewidgetitem5 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(5, __qtablewidgetitem5)
        __qtablewidgetitem6 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(6, __qtablewidgetitem6)
        __qtablewidgetitem7 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(7, __qtablewidgetitem7)
        __qtablewidgetitem8 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(8, __qtablewidgetitem8)
        __qtablewidgetitem9 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(9, __qtablewidgetitem9)
        __qtablewidgetitem10 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(10, __qtablewidgetitem10)
        if (self.tableWidget.rowCount() < 30):
            self.tableWidget.setRowCount(30)
        __qtablewidgetitem11 = QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, __qtablewidgetitem11)
        __qtablewidgetitem12 = QTableWidgetItem()
        __qtablewidgetitem12.setTextAlignment(Qt.AlignCenter);
        self.tableWidget.setItem(0, 0, __qtablewidgetitem12)
        self.tableWidget.setObjectName(u"tableWidget")
        self.tableWidget.setMinimumSize(QSize(500, 300))
        self.tableWidget.setStyleSheet(u"QHeaderView::section {background-color:rgba(162, 255, 134, 100);color: rgb(48, 48, 48);}\n"
"QHeaderView::Stretch")
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.tableWidget.setAlternatingRowColors(False)
        self.tableWidget.setRowCount(30)
        self.tableWidget.horizontalHeader().setVisible(True)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(True)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(80)
        self.tableWidget.horizontalHeader().setHighlightSections(True)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.verticalHeader().setCascadingSectionResizes(False)
        self.tableWidget.verticalHeader().setDefaultSectionSize(30)
        self.tableWidget.verticalHeader().setHighlightSections(False)
        self.tableWidget.verticalHeader().setProperty(u"showSortIndicator", False)

        self.verticalLayout.addWidget(self.tableWidget)

        self.textBrowser = QTextBrowser(BuildTestTableWidget)
        self.textBrowser.setObjectName(u"textBrowser")
        self.textBrowser.setMinimumSize(QSize(0, 30))
        self.textBrowser.setMaximumSize(QSize(16777215, 150))

        self.verticalLayout.addWidget(self.textBrowser)


        self.retranslateUi(BuildTestTableWidget)

        QMetaObject.connectSlotsByName(BuildTestTableWidget)
    # setupUi

    def retranslateUi(self, BuildTestTableWidget):
        BuildTestTableWidget.setWindowTitle(QCoreApplication.translate("BuildTestTableWidget", u"New Procedure", None))
        self.label.setText(QCoreApplication.translate("BuildTestTableWidget", u"Description", None))
        self.plainTextEdit.setPlainText("")
        ___qtablewidgetitem = self.tableWidget.horizontalHeaderItem(0)
        ___qtablewidgetitem.setText(QCoreApplication.translate("BuildTestTableWidget", u"Step", None));
        ___qtablewidgetitem1 = self.tableWidget.horizontalHeaderItem(1)
        ___qtablewidgetitem1.setText(QCoreApplication.translate("BuildTestTableWidget", u"Type", None));
        ___qtablewidgetitem2 = self.tableWidget.horizontalHeaderItem(2)
        ___qtablewidgetitem2.setText(QCoreApplication.translate("BuildTestTableWidget", u"Mode", None));
        ___qtablewidgetitem3 = self.tableWidget.horizontalHeaderItem(3)
        ___qtablewidgetitem3.setText(QCoreApplication.translate("BuildTestTableWidget", u"Value", None));
        ___qtablewidgetitem4 = self.tableWidget.horizontalHeaderItem(4)
        ___qtablewidgetitem4.setText(QCoreApplication.translate("BuildTestTableWidget", u"End Type", None));
        ___qtablewidgetitem5 = self.tableWidget.horizontalHeaderItem(5)
        ___qtablewidgetitem5.setText(QCoreApplication.translate("BuildTestTableWidget", u"Op", None));
        ___qtablewidgetitem6 = self.tableWidget.horizontalHeaderItem(6)
        ___qtablewidgetitem6.setText(QCoreApplication.translate("BuildTestTableWidget", u"Value", None));
        ___qtablewidgetitem7 = self.tableWidget.horizontalHeaderItem(7)
        ___qtablewidgetitem7.setText(QCoreApplication.translate("BuildTestTableWidget", u"Goto", None));
        ___qtablewidgetitem8 = self.tableWidget.horizontalHeaderItem(8)
        ___qtablewidgetitem8.setText(QCoreApplication.translate("BuildTestTableWidget", u"Report Type", None));
        ___qtablewidgetitem9 = self.tableWidget.horizontalHeaderItem(9)
        ___qtablewidgetitem9.setText(QCoreApplication.translate("BuildTestTableWidget", u"Value", None));
        ___qtablewidgetitem10 = self.tableWidget.horizontalHeaderItem(10)
        ___qtablewidgetitem10.setText(QCoreApplication.translate("BuildTestTableWidget", u"Step Note", None));
        ___qtablewidgetitem11 = self.tableWidget.verticalHeaderItem(0)
        ___qtablewidgetitem11.setText(QCoreApplication.translate("BuildTestTableWidget", u"1", None));

        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setSortingEnabled(__sortingEnabled)

        self.textBrowser.setHtml(QCoreApplication.translate("BuildTestTableWidget", u"<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><meta charset=\"utf-8\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"hr { height: 1px; border-width: 0; }\n"
"li.unchecked::marker { content: \"\\2610\"; }\n"
"li.checked::marker { content: \"\\2612\"; }\n"
"</style></head><body style=\" font-family:'\ub9d1\uc740 \uace0\ub515'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"></p></body></html>", None))
    # retranslateUi

