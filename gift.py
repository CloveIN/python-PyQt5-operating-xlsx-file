# !/usr/bin/env python
# encoding: utf-8
# @Time    : 2021/4/7 10:24
# @Author  : Byliner
# @Site    : 
# @File    : gift.py

import sys
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import pandas as pd
import math
import configparser

class GiftPackage(object):

    splitFile = ''  # 打开的文件
    folderName = ''  # 保存的文件夹
    customFile = []  # 自定义分配文件
    isCfg = 1  # 1为基于文件-数量分配；2为基于配置分配
    baseOn = 0  # 0为基于文件数量平均分配；1为基于礼包码数量平均分配
    number = 0

    def setupUi(self, mainWindow):
        mainWindow.setObjectName("mainWindow")
        mainWindow.resize(800, 592)
        self.centralwidget = QtWidgets.QWidget(mainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 81, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.openFile)

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(100, 10, 691, 23))
        self.label.setObjectName("label")

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 50, 83, 23))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.openFolder)

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(98, 50, 691, 23))
        self.label_2.setObjectName("label_2")

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 200, 801, 311))
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setRowCount(0)

        item = QtWidgets.QTableWidgetItem()
        font = QtGui.QFont()
        font.setKerning(True)
        item.setFont(font)
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()

        self.tableWidget.setHorizontalHeaderItem(1, item)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(False)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(300)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(10, 530, 88, 23))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.start)

        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(120, 530, 671, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setObjectName("progressBar")

        self.splitter_2 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_2.setGeometry(QtCore.QRect(10, 90, 249, 19))
        self.splitter_2.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_2.setObjectName("splitter_2")

        self.radioButton = QtWidgets.QRadioButton(self.splitter_2)
        self.radioButton.setObjectName("radioButton")

        self.radioButton_2 = QtWidgets.QRadioButton(self.splitter_2)
        self.radioButton_2.setObjectName("radioButton_2")

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setEnabled(True)
        self.lineEdit.setGeometry(QtCore.QRect(180, 120, 157, 21))
        self.lineEdit.setObjectName("lineEdit")

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setEnabled(True)
        self.label_3.setGeometry(QtCore.QRect(140, 120, 30, 21))
        self.label_3.setObjectName("label_3")

        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setEnabled(True)
        self.comboBox.setGeometry(QtCore.QRect(10, 120, 111, 21))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")

        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(10, 170, 141, 23))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.openSetFile)

        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(360, 119, 75, 23))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.clicked.connect(self.saveConfig)

        mainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(mainWindow)
        self.statusbar.setObjectName("statusbar")
        mainWindow.setStatusBar(self.statusbar)

        #禁止调整大小，禁用最大化按钮
        mainWindow.setFixedSize(mainWindow.width(), mainWindow.height());

        self.retranslateUi(mainWindow)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "礼包工具"))

        self.pushButton.setText(_translate("mainWindow", "源文件"))

        self.label.setText(_translate("mainWindow", "未选择源文件"))

        self.pushButton_2.setText(_translate("mainWindow", "目标文件夹"))

        self.label_2.setText(_translate("mainWindow", "未选择目标文件夹"))

        item = self.tableWidget.horizontalHeaderItem(0)

        item.setText(_translate("mainWindow", "文件名"))

        item = self.tableWidget.horizontalHeaderItem(1)

        item.setText(_translate("mainWindow", "礼包数量"))

        self.pushButton_3.setText(_translate("mainWindow", "开始"))

        self.radioButton.setText(_translate("mainWindow", "按文件名-数量分配"))

        self.radioButton_2.setText(_translate("mainWindow", "按配置分配"))

        self.label_3.setText(_translate("mainWindow", "数量"))

        self.comboBox.setItemText(0, _translate("mainWindow", "文件数量"))

        self.comboBox.setItemText(1, _translate("mainWindow", "礼包数量"))

        self.pushButton_4.setText(_translate("mainWindow", "导入分配配置文件"))

        self.pushButton_5.setText(_translate("mainWindow", "保存配置"))

    def openFile(self):
        _translate = QtCore.QCoreApplication.translate
        fileName, _ = QFileDialog.getOpenFileName(None, '打开excel文件', 'D:\wwwroot', 'Excel files (*.xlsx *.csv)')
        if fileName == "":
            print("\n取消选择")
            return
        self.label.setText(_translate("mainWindow", fileName))
        GiftPackage.splitFile = pd.read_excel(fileName)

    def openSetFile(self):
        _translate = QtCore.QCoreApplication.translate
        fileName, _ = QFileDialog.getOpenFileName(None, '打开txt文件', 'D:\wwwroot', 'TXT files (*.txt)')
        if fileName == "":
            print("\n取消选择")
            return
        with open(fileName, "r", encoding='utf-8') as f:
            lines = f.readlines()
            for i in range(len(lines)):
                fileName, num = lines[i].split('-')
                num = num.strip()
                fileName = fileName.strip()
                self.tableWidget.insertRow(i)
                self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(fileName))
                self.tableWidget.setItem(i, 1, QtWidgets.QTableWidgetItem(num))
                GiftPackage.customFile.append([fileName, num])

    def openFolder(self):
        _translate = QtCore.QCoreApplication.translate
        GiftPackage.folderName = QFileDialog.getExistingDirectory(None, '选择文件夹', 'D:\wwwroot')
        if GiftPackage.folderName == "":
            print("\n取消选择")
            return
        self.label_2.setText(_translate("mainWindow", GiftPackage.folderName))

    def writeConfig(self, mainWindow):
        path = os.path.abspath(os.curdir) + '\config.ini'
        isConfig = os.path.exists(path)
        if not isConfig:
            cf = configparser.ConfigParser()
            cf.add_section('config')
            cf.set('config', 'is_cfg', '1')
            cf.set('config', 'base_on', '0')
            cf.set('config', 'number', '0')
            try:
                with open(path, "a") as f:
                    cf.write(f)
            except ImportError:
                pass

    def initConfigInfo(self, mainWindow):
        cf = configparser.ConfigParser()
        path = os.path.abspath(os.curdir) + '\config.ini'
        cf.read(path)
        a_sections = cf.sections()
        for i in a_sections:
            for j in cf.options(i):
                currentItem = int(cf.get(i, j))
                if j == 'is_cfg':
                    if currentItem == 1:
                        self.radioButton.setChecked(True)
                    else:
                        self.radioButton_2.setChecked(True)
                if j == 'base_on':
                    self.comboBox.setCurrentIndex(currentItem)
                if j == 'number':
                    self.lineEdit.setText(str(currentItem))

    def saveConfig(self):
        path = os.path.abspath(os.curdir) + '\config.ini'
        #获取配置
        if self.radioButton.isChecked():
            is_cfg = 1
        else:
            is_cfg = 2
        base_on = self.comboBox.currentIndex()
        number = self.lineEdit.text()
        cf = configparser.ConfigParser()
        cf.read(path)
        cf.set('config', 'is_cfg', str(is_cfg))
        cf.set('config', 'base_on', str(base_on))
        cf.set('config', 'number', number)
        try:
            with open(path, "w+") as f:
                cf.write(f)
        except ImportError:
            print('保存错误')
            self.statusbar.showMessage('保存错误')
            pass
        self.statusbar.showMessage('保存成功')


    def start(self):
        _translate = QtCore.QCoreApplication.translate
        df = GiftPackage.splitFile
        if type(df) == str:
            self.statusbar.showMessage('请选择待分配文件')
            return
        rows, cols = df.shape
        if self.radioButton.isChecked():
            is_cfg = 1  #基于文件名-数量分配
        else:
            is_cfg = 2  #基于配置分配
        base_on = self.comboBox.currentIndex()
        number = self.lineEdit.text()

        doneProgress = 0 #完成进度
        allCount = 0 #总量
        if is_cfg == 1:
            fileArr = GiftPackage.customFile
            k = 0
            new_list = []
            rows_format = 0
            if not fileArr:
                self.statusbar.showMessage('请选择分配配置文件')
                return
            for fileItem in fileArr:
                fileName,num = fileItem
                num = int(num)
                if (k + num) >= rows:
                    break
                new_list.append([fileName, k, k + num])
                k = k + num
                rows_format = k
            allCount = len(new_list)
            k = 0
            for f_i_j in new_list:
                self.pushButton_3.setEnabled(False)
                self.pushButton_3.setText(_translate("mainWindow", "停止"))
                k = k + 1
                doneProgress = (k / allCount) * 100
                f, i, j = f_i_j
                excel_small = df[i:j]
                excel_small.to_excel('{0}_{1}_{2}.xlsx'.format(f, i, j), index=False)
                if rows > rows_format:
                    df[rows_format:].to_excel('礼包文件_last.xlsx')
                self.progressBar.setValue(int(doneProgress))
        else:
            if base_on == 0:
                # 基于文件数量平均分配
                split_num = math.floor(rows / int(number))
            else:
                # 基于礼包数量平均分配
                split_num = int(number)
            value = math.floor(rows / split_num)
            rows_format = value * split_num
            new_list = [[i, i + split_num] for i in range(0, rows_format, split_num)]
            allCount = len(new_list)
            k = 0
            for i_j in new_list:
                self.pushButton_3.setEnabled(False)
                self.pushButton_3.setText(_translate("mainWindow", "停止"))
                k = k + 1
                doneProgress = (k / allCount) * 100
                i, j = i_j
                excel_small = df[i:j]
                excel_small.to_excel('礼包文件_{0}_{1}.xlsx'.format(i, j), index=False)
                if rows > rows_format:
                    df[rows_format:].to_excel('礼包文件_last.xlsx')
                self.progressBar.setValue(int(doneProgress))
        if doneProgress >= 100:
            self.pushButton_3.setEnabled(True)
            self.pushButton_3.setText(_translate("mainWindow", "开始"))
        self.statusbar.showMessage('分配完成')


def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = GiftPackage()
    ui.setupUi(MainWindow)
    ui.writeConfig(MainWindow)
    ui.initConfigInfo(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()