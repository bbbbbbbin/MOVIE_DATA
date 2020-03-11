# -*- coding:utf-8 -*-

import xlrd
import re
import xlwt
from xlutils.copy import copy

book = xlrd.open_workbook('data/2017baidu.xls')
sheet = book.sheet_by_index(0)
nb = copy(book)
nw = nb.get_sheet(0)
names = sheet.col_values(1)
for i in range(len(names)):
    k = 0
    all = 0
    for j in range(30):
        if sheet.col_values(3+j)[i] != -1:
            k += 1
            print sheet.col_values(3+j)[i]
            all += float(sheet.col_values(3+j)[i])
    if all == 0:
        nw.write(i,2,0)
    else:
        nw.write(i,2,all/k)
nb.save('data/2017baidu.xls')