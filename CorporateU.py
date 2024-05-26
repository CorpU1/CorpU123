

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from CorporateU import Ui_MainWindow  # Импортируем сгенерированный файл
from docx import Document

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(812, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.formLayout = QtWidgets.QFormLayout(self.centralwidget)
        self.formLayout.setObjectName("formLayout")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setObjectName("label_3")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_3)
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setObjectName("textEdit_2")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.textEdit_2)
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setObjectName("label_6")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.label_6)
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setObjectName("textEdit_3")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.textEdit_3)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.label)
        self.dateEdit = QtWidgets.QDateEdit(self.centralwidget)
        self.dateEdit.setObjectName("dateEdit")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.dateEdit)
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setObjectName("label_7")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.LabelRole, self.label_7)
        self.textEdit_6 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_6.setObjectName("textEdit_6")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.FieldRole, self.textEdit_6)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setObjectName("label_5")
        self.formLayout.setWidget(7, QtWidgets.QFormLayout.LabelRole, self.label_5)
        self.spinBox_2 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_2.setObjectName("spinBox_2")
        self.formLayout.setWidget(7, QtWidgets.QFormLayout.FieldRole, self.spinBox_2)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setObjectName("label_4")
        self.formLayout.setWidget(10, QtWidgets.QFormLayout.LabelRole, self.label_4)
        self.spinBox = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox.setObjectName("spinBox")
        self.formLayout.setWidget(10, QtWidgets.QFormLayout.FieldRole, self.spinBox)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.formLayout.setLayout(12, QtWidgets.QFormLayout.SpanningRole, self.horizontalLayout)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 812, 18))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menu.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_3.setText(_translate("MainWindow", "Цели"))
        self.label_6.setText(_translate("MainWindow", "Описание бизнеса"))
        self.label.setText(_translate("MainWindow", "Дата"))
        self.label_7.setText(_translate("MainWindow", "Действия, принятые на первом  ОСУ"))
        self.label_5.setText(_translate("MainWindow", "Доля Участника 1"))
        self.label_4.setText(_translate("MainWindow", "Количество СД"))
        self.pushButton_2.setText(_translate("MainWindow", "ОК"))
        self.menu.setTitle(_translate("MainWindow", "КорпСтудия"))


class CorporateApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.generate_contract)

    def generate_contract(self):
        sd = self.spinBox.value()
        stock = self.spinBox_2.value()
        stock2 = 100 - sd
        decisions = self.textEdit_6.toPlainText()
        business = self.textEdit_3.toPlainText()
        target = self.textEdit_2.toPlainText()

        # Путь для сохранения сгенерированного договора
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить договор", "",
                                                             "Word Documents (*.docx);;All Files (*)", options=options)

        if file_path:
            self.create_contract(sd, stock, stock2, decisions, business, target, file_path)
            QtWidgets.QMessageBox.information(self, "Успех", "Договор успешно создан!")

    def create_contract(self, sd, stock, stock2, decisions, business, target, output_path):
        doc = Document('contract_template.docx')
        for p in doc.paragraphs:
            if '{SD}' in p.text:
                p.text = p.text.replace('{SD}', str(sd))
            if '{stock}' in p.text:
                p.text = p.text.replace('{stock}', str(stock))
            if '{stock2}' in p.text:
                p.text = p.text.replace('{stock2}', str(stock2))
            if '{decitions}' in p.text:
                p.text = p.text.replace('{decitions}', decisions)
            if '{business}' in p.text:
                p.text = p.text.replace('{business}', business)
            if '{target}' in p.text:
                p.text = p.text.replace('{target}', target)
        doc.save(output_path)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())