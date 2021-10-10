#!/usr/bin/python
# -*- coding: utf-8 -*-
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from main import *
import sys
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def get_date(data_0):
    date_0 = data_0['日期']
    date_1 = date_0.dropna()
    date_2 = date_1.tolist()
    for i in range(len(date_2)):
        date_2[i] = str(date_2[i].date())
    return date_2


class MyWindow(QMainWindow, Ui_mainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.reads)
        self.comboBox_2.activated.connect(self.get_date)
        self.comboBox_3.activated.connect(self.maintenance)
        self.file_name = ''
        self.file_sheet = ''
        self.date_get = ''
        self.data = pd.DataFrame()
        self.date = []
        self.path = ''

    # 读取文件
    def reads(self):
        file_name0 = QFileDialog.getOpenFileName(self, '选择读取文件', filter='EXCEL FILE(*.xlsx )')
        self.file_name = file_name0[0]
        if self.file_name != '':
            # 读取文件地址
            path_s = self.file_name.split('/')
            separator = '/'
            self.path = separator.join(path_s[:-1])
            # 判断文件合法性
            suffix = path_s[-1].split(',')
            self.readTabParams()

    def readTabParams(self):
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
    @staticmethod
    def location(data, date_list, date_get):
        # 所选日期在表格内位置
        date_position1 = [data[data['日期'] == date_get].index.tolist()[0]]
        # 输入日期所在列表位置
        p = date_list.index(str(date_get))
        # 最后一日数据
        if len(date_list) - p <= 2:
            data_class = data.iloc[date_position1[0]:, 1]
            calls_form = data_class[data_class.str.contains('班', na=False)]
            calls_position = calls_form.index.tolist()
            date_position_re = [calls_position[0], data.shape[0]]
        # 非最后一日数据
        else:
            date_position1.append(data[data['日期'] == date_list[p + 2]].index.tolist()[0])
            data_class = data.iloc[date_position1[0]:date_position1[1], 1]
            calls_form = data_class[data_class.str.contains('班', na=False)]
            calls_position = calls_form.index.tolist()
            date_position_re = [calls_position[0], calls_position[4]]
        # 第二列中
        return date_position_re

    # 获得去重数据
    @staticmethod
    def duplicates(data):
        a = data.columns.values
        data = data.drop_duplicates(subset=[a[2], a[6], a[9], a[10]], keep='first')
        return data

    # 求和
    @staticmethod
    def SumDate(DataInstance):
        summation = DataInstance.iloc[0, 2:-2]

        if DataInstance.shape[0] > 1:
            summation.iloc[1] = float(DataInstance.iloc[:, 3].sum())
            summation.iloc[2] = float(DataInstance.iloc[:, 4].sum())
            summation.iloc[3] = float(DataInstance.iloc[:, 5].sum())
            summation.iloc[5] = float(DataInstance.iloc[:, 7].sum())
            summation.iloc[6] = float(DataInstance.iloc[:, 8].sum())
            summation.iloc[9] = float(DataInstance.iloc[:, 11].sum())
            summation.iloc[10] = float(DataInstance.iloc[:, 12].sum())
            summation.iloc[12] = float(DataInstance.iloc[:, 14].sum())
            summation.iloc[14] = float(DataInstance.iloc[:, 16].sum())
            summation.iloc[16] = float(DataInstance.iloc[:, 18].sum())
            summation.iloc[18] = float(DataInstance.iloc[:, 20].sum())
            summation.iloc[20] = float(DataInstance.iloc[:, 22].sum())
            summation.iloc[22] = float(DataInstance.iloc[:, 24].sum())
            summation.iloc[24] = float(DataInstance.iloc[:, 26].sum())
            summation.iloc[26] = float(DataInstance.iloc[:, 28].sum())
            summation.iloc[7] = float(summation.iloc[9]) / float(summation.iloc[6]) * 100
            summation.iloc[8] = float(summation.iloc[10]) / float(summation.iloc[6]) * 100
            summation.iloc[11] = float(summation.iloc[12]) / float(summation.iloc[6]) * 100
            summation.iloc[13] = float(summation.iloc[14]) / float(summation.iloc[6]) * 100
            summation.iloc[15] = float(summation.iloc[16]) / float(summation.iloc[6]) * 100
            summation.iloc[17] = float(summation.iloc[18]) / float(summation.iloc[6]) * 100
            summation.iloc[19] = float(summation.iloc[20]) / float(summation.iloc[6]) * 100
            summation.iloc[21] = float(summation.iloc[22]) / float(summation.iloc[6]) * 100
            summation.iloc[23] = float(summation.iloc[24]) / float(summation.iloc[6]) * 100
            summation.iloc[25] = float(summation.iloc[26]) / float(summation.iloc[6]) * 100
        return summation

    # 处理数据
    def createProcess(self, data_clean, data_duplicates):
        final_list = []
        for i in range(data_duplicates.shape[0]):
            equipment = data_duplicates.iloc[i, 2]
            loading_place = data_duplicates.iloc[i, 6]
            copper = data_duplicates.iloc[i, 9]
            molybdenum = data_duplicates.iloc[i, 10]
            CompareAction = [equipment, loading_place, copper, molybdenum]
            DataInstance = data_clean[
                (data_clean.iloc[:, 2] == CompareAction[0]) & (data_clean.iloc[:, 6] == CompareAction[1]) & (
                        data_clean.iloc[:, 9] == CompareAction[2]) & (data_clean.iloc[:, 10] == CompareAction[3])]
            summation = self.SumDate(DataInstance)
            summation = summation.values.tolist()
            final_list.append(summation)
        final = pd.DataFrame(final_list, columns=data_duplicates.columns.values[2:-2])
        return final

    # 输出数据
    def output(self, summat, date_get):
        name = self.path + '/' + str(date_get) + '报表.xlsx'
        summat = summat.sort_values(by='装矿设备')
        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(summat, index=False, header=True):
            ws.append(r)
        wb.save(name)
        self.label_2.setText('输出成功')

    def maintenance(self):
        file_name = self.file_name
        data = self.data
        date_get = self.comboBox_3.currentText()
        date_list = self.date
        date_position = self.location(data, date_list, date_get)
        # 选取所需数据
        data_slice = data.iloc[date_position[0]: date_position[1]]
        # 清洗数据
        data_clean = data_slice.dropna(subset=['装矿设备'])
        # 去重数据表
        data_duplicates = self.duplicates(data_clean)
        # 求和
        summat = self.createProcess(data_clean, data_duplicates)
        self.output(summat, date_get)


app = QApplication(sys.argv)
dc = MyWindow()
dc.show()
sys.exit(app.exec_())
