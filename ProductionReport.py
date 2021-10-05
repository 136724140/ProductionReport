#!/usr/bin/python
# -*- coding: utf-8 -*-
import pandas as pd

file_name = '2021年一二系列日供矿记录表(新).xls'
file_sheet = '九月份'
date_get = '2021-09-26'


# 获取日期列表
def get_date(data_0):
    date_0 = data_0['日期']
    date_1 = date_0.dropna()
    date_2 = date_1.tolist()
    for i in range(len(date_2)):
        date_2[i] = str(date_2[i].date())
    return date_2


# 获取所选日期位置范围
def positon(data, date_list, date_get):
    date_position1 = [data[data['日期'] == date_get].index.tolist()[0]]
    p = date_list.index(str(date_get))
    if len(date_list) - date_list.index(str(date_get)) <= 3:
        data_class = data.iloc[date_position1[0]:, 1]
    else:
        date_position1.append(data[data['日期'] == date_list[p + 2]].index.tolist()[0])
        data_class = data.iloc[date_position1[0]:date_position1[1] + 1, 1]
    # 第二列中
    calls_form = data_class[data_class.str.contains('班', na=False)]
    calls_position = calls_form.index.tolist()
    date_position_re = [calls_position[0], calls_position[4] - 1]
    return date_position_re


# 获取sheet名
df = pd.read_excel(file_name, sheet_name=None)
# 读取指定sheet文件
data = pd.read_excel(file_name, sheet_name=file_sheet, header=1)
date_list = get_date(data)
date_position = positon(data, date_list, date_get)
# 选取所需数据
date_slice = data.iloc[date_position[0]: date_position[1]]
print(date_slice)
# data_select = data.iloc[date_position[0]:date_position[1]]
# print(data_select)
