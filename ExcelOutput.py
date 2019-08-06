# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'g_project.ui'
#
# Created by: PyQt5 UI code generator 5.12.3
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import pandas as pd
import numpy as np
import csv
import xlwt
import os

class Ui_MainWindow(QWidget):
    global in_count
    in_count = 0
    global path_openfile_name
    path_openfile_name = " "
    def openfile(self):
        fileName = QFileDialog.getOpenFileName(self, "Open Excel", "./", "Excel files(*.xlsx , *.xls)")
        global path_openfile_name
        path_openfile_name = fileName[0]
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", path_openfile_name))

    def filesave(self):
        global path_openfile_name
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet")
        self.add2(sheet)

        if path_openfile_name == " ":
            Name = QFileDialog.getSaveFileName(self, "Save file", "./", "xls(*.xls)")
            # print(Name)
            wbk.save(Name[0])
            _translate = QtCore.QCoreApplication.translate
            MainWindow.setWindowTitle(_translate("MainWindow", Name[0]))
        else:
            wbk.save(path_openfile_name)

    def filesaveas(self):
        global path_openfile_name
        Name = QFileDialog.getSaveFileName(self, "Save file", "./", "xls(*.xls)")
        # print(Name)
        path_openfile_name = Name[0]
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", path_openfile_name))
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet")
        self.add2(sheet)
        wbk.save(Name[0])

    def add2(self, sheet):
        horizontalHeader = ["Name", "Platform", "Type", "Status", "Remarks"]
        for currentColumn in range(5):
            sheet.write(0, currentColumn, horizontalHeader[currentColumn])

        for currentRow in range(self.tableWidget.rowCount()):
            for currentColumn in range(self.tableWidget.columnCount()):
                listdata = str(self.tableWidget.item(currentRow, currentColumn).text())
                sheet.write(currentRow + 1, currentColumn, listdata)

    def finishedEdit(self):
        self.tableWidget.resizeColumnsToContents()

    def creat_table_show(self):
        global in_count
        if len(path_openfile_name) > 0:
            input_table = pd.read_excel(path_openfile_name)
            # print(input_table)
            input_table_rows = input_table.shape[0]
            input_table_colunms = input_table.shape[1]
            # print(input_table_rows)
            # print(input_table_colunms)
            input_table_header = input_table.columns.values.tolist()
            # print(input_table_header)

            self.tableWidget.setColumnCount(input_table_colunms)
            self.tableWidget.setRowCount(input_table_rows)
            self.tableWidget.setHorizontalHeaderLabels(input_table_header)

            for i in range(input_table_rows):
                input_table_rows_values = input_table.iloc[[i]]
                # print(input_table_rows_values)
                input_table_rows_values_array = np.array(input_table_rows_values)
                input_table_rows_values_list = input_table_rows_values_array.tolist()[0]
                print(input_table_rows_values_list)
                for j in range(input_table_colunms):
                    input_table_items_list = input_table_rows_values_list[j]
                    # print(input_table_items_list)
                    # print(type(input_table_items_list))

                    input_table_items = str(input_table_items_list)
                    newItem = QTableWidgetItem(input_table_items)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget.setItem(i, j, newItem)

            in_count = input_table_rows

        # else:
            # self.centralWidget.show()

    def add_data(self):
        global g_name
        g_name = self.name_input.toPlainText() + "," + self.index_list.currentText() + "," + self.format_list.currentText() + "," \
                 + self.status_list.currentText() + "," + self.ex_input.toPlainText()
        print (g_name)

    def add_show(self):
        global in_count
        if in_count == 0:
            horizontalHeader = ["Name","Platform","Type","Status","Remark"]
            self.tableWidget.setColumnCount(5)
            self.tableWidget.setRowCount(1)
            self.tableWidget.setHorizontalHeaderLabels(horizontalHeader)
            # print (len(g_name))
            g_name_show = g_name.split(',')
            # print (g_name_show)
            for j in range(5):
                self.tableWidget.setItem(0, j, QTableWidgetItem(g_name_show[j]))
        else :
            inrow = in_count+1
            self.tableWidget.setRowCount(inrow)
            g_name_show = g_name.split(',')
            for j in range(5):
                self.tableWidget.setItem(in_count, j, QTableWidgetItem(g_name_show[j]))
            print("COUNT = "+str(in_count+1))
        in_count += 1

    def clear_data(self):
        self.name_input.clear()
        self.index_list.setCurrentIndex(0)
        self.format_list.setCurrentIndex(0)
        self.status_list.setCurrentIndex(0)
        self.ex_input.clear()

    def ver_event(self):
        QMessageBox.about(self, 'About', 'Designer：Leon Lee\nVersion：V1.0')

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(931, 415)
        font = QtGui.QFont()
        font.setKerning(True)
        MainWindow.setFont(font)
        MainWindow.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        MainWindow.setMouseTracking(False)
        MainWindow.setTabletTracking(False)
        MainWindow.setLayoutDirection(QtCore.Qt.LeftToRight)
        MainWindow.setAutoFillBackground(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.index = QtWidgets.QLabel(self.centralwidget)
        self.index.setGeometry(QtCore.QRect(260, 0, 100, 31))
        font = QtGui.QFont()
        font.setFamily("Apple Color Emoji")
        font.setPointSize(18)
        self.index.setFont(font)
        self.index.setTextFormat(QtCore.Qt.AutoText)
        self.index.setAlignment(QtCore.Qt.AlignCenter)
        self.index.setObjectName("index")

        self.index_list = QtWidgets.QComboBox(self.centralwidget)
        self.index_list.setGeometry(QtCore.QRect(240, 20, 135, 51))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.index_list.setFont(font)
        self.index_list.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.index_list.setObjectName("index_list")
        self.index_list.addItem("")
        self.index_list.addItem("")
        self.index_list.addItem("")

        self.name = QtWidgets.QLabel(self.centralwidget)
        self.name.setGeometry(QtCore.QRect(80, 0, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Apple Color Emoji")
        font.setPointSize(18)
        self.name.setFont(font)
        self.name.setTextFormat(QtCore.Qt.AutoText)
        self.name.setAlignment(QtCore.Qt.AlignCenter)
        self.name.setObjectName("name")

        self.name_input = QtWidgets.QTextEdit(self.centralwidget)
        self.name_input.setGeometry(QtCore.QRect(30, 30, 181, 31))
        self.name_input.setObjectName("name_input")

        self.status = QtWidgets.QLabel(self.centralwidget)
        self.status.setGeometry(QtCore.QRect(550, 0, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Apple Color Emoji")
        font.setPointSize(18)
        self.status.setFont(font)
        self.status.setTextFormat(QtCore.Qt.AutoText)
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        self.status.setObjectName("status")
        self.status_list = QtWidgets.QComboBox(self.centralwidget)
        self.status_list.setGeometry(QtCore.QRect(530, 20, 111, 51))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.status_list.setFont(font)
        self.status_list.setToolTipDuration(-1)
        self.status_list.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.status_list.setObjectName("status_list")
        self.status_list.addItem("")
        self.status_list.addItem("")
        self.status_list.addItem("")
        self.format = QtWidgets.QLabel(self.centralwidget)
        self.format.setGeometry(QtCore.QRect(410, 0, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Apple Color Emoji")
        font.setPointSize(18)
        self.format.setFont(font)
        self.format.setTextFormat(QtCore.Qt.AutoText)
        self.format.setAlignment(QtCore.Qt.AlignCenter)
        self.format.setObjectName("format")
        self.format_list = QtWidgets.QComboBox(self.centralwidget)
        self.format_list.setGeometry(QtCore.QRect(390, 20, 111, 51))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.format_list.setFont(font)
        self.format_list.setToolTipDuration(-1)
        self.format_list.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.format_list.setObjectName("format_list")
        self.format_list.addItem("")
        self.format_list.addItem("")
        self.ex = QtWidgets.QLabel(self.centralwidget)
        self.ex.setGeometry(QtCore.QRect(710, 0, 71, 31))
        font = QtGui.QFont()
        font.setFamily("Apple Color Emoji")
        font.setPointSize(18)
        self.ex.setFont(font)
        self.ex.setTextFormat(QtCore.Qt.AutoText)
        self.ex.setAlignment(QtCore.Qt.AlignCenter)
        self.ex.setObjectName("ex")
        self.ex_input = QtWidgets.QTextEdit(self.centralwidget)
        self.ex_input.setGeometry(QtCore.QRect(660, 30, 181, 31))
        self.ex_input.setObjectName("ex_input")

        self.add_list = QtWidgets.QPushButton(self.centralwidget)
        self.add_list.setGeometry(QtCore.QRect(350, 80, 121, 41))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.add_list.setFont(font)
        self.add_list.setObjectName("add_list")
        self.add_list.clicked.connect(self.add_data)
        self.add_list.clicked.connect(self.add_show)

        self.clear = QtWidgets.QPushButton(self.centralwidget)
        self.clear.setGeometry(QtCore.QRect(480, 80, 121, 41))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.clear.setFont(font)
        self.clear.setObjectName("clear")
        self.clear.clicked.connect(self.clear_data)

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 130, 931, 251))
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setObjectName("tableWidget")
        # self.model.dataChanged.connect(self.finishedEdit)
        self.tableWidget.raise_()


        MainWindow.setCentralWidget(self.centralwidget)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 931, 22))
        self.menuBar.setObjectName("menuBar")
        self.menu = QtWidgets.QMenu(self.menuBar)
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menuBar)
        self.menu_2.setObjectName("menu_2")
        MainWindow.setMenuBar(self.menuBar)
        self.toolBar = QtWidgets.QToolBar(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(16)
        self.toolBar.setFont(font)
        self.toolBar.setObjectName("toolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)

        self.file_open = QtWidgets.QAction(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(13)
        self.file_open.setFont(font)
        self.file_open.setObjectName("file_open")
        self.file_open.triggered.connect(self.openfile)
        self.file_open.triggered.connect(self.creat_table_show)

        self.file_save = QtWidgets.QAction(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(13)
        self.file_save.setFont(font)
        self.file_save.setObjectName("file_save")
        self.file_save.triggered.connect(self.filesave)

        self.file_save_as = QtWidgets.QAction(MainWindow)
        self.file_save_as.setObjectName("file_save_as")
        self.file_save_as.triggered.connect(self.filesaveas)

        self.version = QtWidgets.QAction(MainWindow)
        self.version.setObjectName("version")
        self.version.triggered.connect(self.ver_event)

        self.menu.addAction(self.file_open)
        self.menu.addAction(self.file_save)
        self.menu.addAction(self.file_save_as)
        self.menu_2.addAction(self.version)
        self.menuBar.addAction(self.menu.menuAction())
        self.menuBar.addAction(self.menu_2.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ExcelOutput"))
        self.index.setText(_translate("MainWindow", "<html><head/><body><p>Platform</p></body></html>"))
        self.index_list.setItemText(0, _translate("MainWindow", "PS4"))
        self.index_list.setItemText(1, _translate("MainWindow", "Steam"))
        self.index_list.setItemText(2, _translate("MainWindow", "NS"))
        self.name.setText(_translate("MainWindow", "<html><head/><body><p>Name</p></body></html>"))
        self.status.setText(_translate("MainWindow", "<html><head/><body><p>Status</p></body></html>"))
        self.status_list.setItemText(0, _translate("MainWindow", "Action"))
        self.status_list.setItemText(1, _translate("MainWindow", "Idle"))
        self.status_list.setItemText(2, _translate("MainWindow", "End"))
        self.format.setText(_translate("MainWindow", "<html><head/><body><p>Type</p></body></html>"))
        self.format_list.setItemText(0, _translate("MainWindow", "Disk"))
        self.format_list.setItemText(1, _translate("MainWindow", "Digits"))
        self.ex.setText(_translate("MainWindow", "<html><head/><body><p>Remark</p></body></html>"))
        self.add_list.setText(_translate("MainWindow", "Add"))
        self.clear.setText(_translate("MainWindow", "Clear"))
        self.menu.setTitle(_translate("MainWindow", "File"))
        self.menu_2.setTitle(_translate("MainWindow", "About"))
        self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        self.file_open.setText(_translate("MainWindow", "Open"))
        self.file_open.setToolTip(_translate("MainWindow", "open"))
        self.file_save.setText(_translate("MainWindow", "Save"))
        self.file_save.setToolTip(_translate("MainWindow", "save"))
        self.file_save_as.setText(_translate("MainWindow", "Save as"))
        self.version.setText(_translate("MainWindow", "Programmer"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
