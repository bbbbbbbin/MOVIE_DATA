# -*- coding:utf -8-*-

import xlrd
import sys
from xlutils.copy import copy
reload(sys)
sys.setdefaultencoding('utf8')

path = 'data/'

def load_dic():
    dic = {}
    book = xlrd.open_workbook(path+'produce.xls')
    sheet = book.sheet_by_index(0)
    coun = sheet.row_values(0)[1:]
    for i in range(len(coun)):
        value = sheet.col_values(i+1)[1] + sheet.col_values(i+1)[3] + \
                sheet.col_values(i+1)[5] + sheet.col_values(i+1)[7] + \
                sheet.col_values(i+1)[9] + sheet.col_values(i+1)[11]
        num = sheet.col_values(i+1)[2] + sheet.col_values(i+1)[4] + \
            sheet.col_values(i+1)[6] + sheet.col_values(i+1)[8] + \
            sheet.col_values(i+1)[10] + sheet.col_values(i+1)[12]
        dic[coun[i]] = value/num
    return dic

dic = load_dic()
for key,word in dic.items():
    print key,' ',word
book = xlrd.open_workbook(path+'2017.xls')
sheet = book.sheet_by_index(0)
newb = copy(book)
news = newb.get_sheet(0)
couns = sheet.col_values(13)
i = 0
for coun in couns:
    v = 0
    cs = coun.split('/')
    for c in cs:
        if c in dic:
            if dic[c] > v:
                v = dic[c]
    news.write(i,13,v)
    i+=1
newb.save(path+'2017.xls')