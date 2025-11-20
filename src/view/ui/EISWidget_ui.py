# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'EISset_Widget.ui'
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
from PySide6.QtWidgets import (QApplication, QComboBox, QLabel, QLineEdit,
    QPushButton, QSizePolicy, QWidget)

class EISUI(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"EIS_Paremeter")
        Form.resize(180, 300)
        self.pushButton = QPushButton(Form)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(10, 10, 75, 24))
        self.pushButton_2 = QPushButton(Form)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(90, 10, 75, 24))
        self.comboBox = QComboBox(Form)
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(QRect(50, 60, 69, 22))
        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(60, 40, 48, 16))
        self.label.setAlignment(Qt.AlignCenter)
        self.lineEdit = QLineEdit(Form)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(30, 110, 113, 21))
        self.lineEdit.setAlignment(Qt.AlignCenter)
        self.label_2 = QLabel(Form)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(40, 90, 91, 16))
        self.label_2.setAlignment(Qt.AlignCenter)
        self.lineEdit_2 = QLineEdit(Form)
        self.lineEdit_2.setObjectName(u"lineEdit_2")
        self.lineEdit_2.setGeometry(QRect(30, 160, 113, 21))
        self.lineEdit_2.setAlignment(Qt.AlignCenter)
        self.label_3 = QLabel(Form)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(40, 140, 91, 16))
        self.label_3.setAlignment(Qt.AlignCenter)
        self.lineEdit_3 = QLineEdit(Form)
        self.lineEdit_3.setObjectName(u"lineEdit_3")
        self.lineEdit_3.setGeometry(QRect(30, 210, 113, 21))
        self.lineEdit_3.setAlignment(Qt.AlignCenter)
        self.label_4 = QLabel(Form)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(50, 190, 71, 16))
        self.label_4.setAlignment(Qt.AlignCenter)
        self.lineEdit_4 = QLineEdit(Form)
        self.lineEdit_4.setObjectName(u"lineEdit_4")
        self.lineEdit_4.setGeometry(QRect(30, 260, 113, 21))
        self.lineEdit_4.setAlignment(Qt.AlignCenter)
        self.label_5 = QLabel(Form)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(40, 240, 91, 16))
        self.label_5.setAlignment(Qt.AlignCenter)
        self.label_6 = QLabel(Form)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(140, 110, 21, 16))
        self.label_6.setAlignment(Qt.AlignCenter)
        self.label_7 = QLabel(Form)
        self.label_7.setObjectName(u"label_7")
        self.label_7.setGeometry(QRect(140, 160, 21, 16))
        self.label_7.setAlignment(Qt.AlignCenter)
        self.label_8 = QLabel(Form)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setGeometry(QRect(140, 210, 21, 16))
        self.label_8.setAlignment(Qt.AlignCenter)

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.pushButton.setText(QCoreApplication.translate("Form", u"\ubd88\ub7ec\uc624\uae30", None))
        self.pushButton_2.setText(QCoreApplication.translate("Form", u"\uc800\uc7a5", None))
        self.comboBox.setItemText(0, QCoreApplication.translate("Form", u"GEIS", None))
        self.comboBox.setItemText(1, QCoreApplication.translate("Form", u"PEIS", None))
        self.comboBox.setItemText(2, QCoreApplication.translate("Form", u"ACIR", None))

        self.label.setText(QCoreApplication.translate("Form", u"Mode", None))
        self.lineEdit.setText(QCoreApplication.translate("Form", u"100000.000", None))
        self.label_2.setText(QCoreApplication.translate("Form", u"Start_Frequency", None))
        self.lineEdit_2.setText(QCoreApplication.translate("Form", u"0.100", None))
        self.label_3.setText(QCoreApplication.translate("Form", u"Stop_Frequency", None))
        self.lineEdit_3.setText(QCoreApplication.translate("Form", u"0.010000", None))
        self.label_4.setText(QCoreApplication.translate("Form", u"Amplitude", None))
        self.lineEdit_4.setText(QCoreApplication.translate("Form", u"50", None))
        self.label_5.setText(QCoreApplication.translate("Form", u"PointNumber", None))
        self.label_6.setText(QCoreApplication.translate("Form", u"Hz", None))
        self.label_7.setText(QCoreApplication.translate("Form", u"Hz", None))
        self.label_8.setText(QCoreApplication.translate("Form", u"A", None))
    # retranslateUi

