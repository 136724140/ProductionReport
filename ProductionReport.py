#!/usr/bin/python
# -*- coding: utf-8 -*-
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QDialog
from main import *
import sys


def get_date(data_0):
    date_0 = data_0['日期']
    date_1 = date_0.dropna()
    date_2 = date_1.tolist()
    for i in range(len(date_2)):
        date_2[i] = str(date_2[i].date())
    print(date_2)
    return date_2


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.readfiel)
        self.comboBox_2.activated.connect(self.get_date)
        self.comboBox_3.activated.connect(self.maintenance)
        self.file_name = '2021年一二系列日供矿记录表(新).xls'
        self.file_sheet = '九月份'
        self.date_get = '2021-09-26'

    # 读取文件
    def readfiel(self):
        file_name0 = QFileDialog.getOpenFileName(self, '选择读取文件')
        self.file_name = file_name0[0]
        self.readsheet()
        # 显示所选文件名

    def readsheet(self):
        file_name = self.file_name
        self.label.setText(file_name)
        # 获取sheet名
        sheet_list = list(pd.read_excel(file_name, sheet_name=None))
        # sheet映射至下拉菜单
        self.comboBox_2.addItems(sheet_list)
        self.file_name = file_name

    def get_date(self):
        self.comboBox_3.clear()
        file_name = self.file_name
        file_sheet = self.comboBox_2.currentText()
        # 读取指定sheet文件
        self.data = pd.read_excel(file_name, sheet_name=file_sheet, header=1)
        data = self.data
        date_0 = data['日期']
        date_1 = date_0.dropna()
        date_2 = date_1.tolist()
        for i in range(len(date_2)):
            date_2[i] = str(date_2[i].date())
        self.date = date_2
        self.comboBox_3.addItems(date_2)


    # 获取所选日期位置范围
    def positon(self, data, date_list, date_get):
        # 所选日期在表格内位置
        date_position1 = [data[data['日期'] == date_get].index.tolist()[0]]
        # 输入日期所在列表位置
        p = date_list.index(str(date_get))
        # 最后一天选到最后一行
        print(len(date_list))
        print(p)

        if len(date_list) - p <= 2:
            data_class = data.iloc[date_position1[0]:, 1]
            calls_form = data_class[data_class.str.contains('班', na=False)]
            calls_position = calls_form.index.tolist()
            date_position_re = [calls_position[0], data.shape[0]-2]
        else:
            date_position1.append(data[data['日期'] == date_list[p + 2]].index.tolist()[0])
            data_class = data.iloc[date_position1[0]:date_position1[1], 1]
            calls_form = data_class[data_class.str.contains('班', na=False)]
            calls_position = calls_form.index.tolist()
            date_position_re = [calls_position[0], calls_position[4]-1]
        print(date_position_re)
        # 第二列中
        return date_position_re

    def maintenance(self):
        file_name = self.file_name
        data = self.data
        date_get = self.comboBox_3.currentText()
        date_list = self.date
        date_position = self.positon(data, date_list, date_get)
        # 选取所需数据
        date_slice = data.iloc[date_position[0]: date_position[1]]
        # data_select = data.iloc[date_position[0]:date_position[1]]
        # print(data_select)


app = QApplication(sys.argv)
mywin = MyWindow()
mywin.show()
sys.exit(app.exec_())
