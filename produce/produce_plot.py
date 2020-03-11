# -*- coding:utf -8-*-

import xlrd
import sys
reload(sys)
import matplotlib.pyplot as plt
sys.setdefaultencoding('utf8')
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示坐标轴负号

xlsfile = r'data/'
picfile = r'pic/'

book = xlrd.open_workbook(xlsfile+'produce.xls')
sheet = book.sheet_by_index(0)
clen = len(sheet.row_values(0))-1
year = [2010,2011,2012,2013,2014,2015]
genre = sheet.row_values(0)[1:]

def Year():
    for i in range(clen):
        value = []
        for j in range(1,13,2):
            if sheet.col_values(i+1)[j+1] == 0:
                value.append(0)
            else:
                value.append(sheet.col_values(i+1)[j] / sheet.col_values(i+1)[j+1])
        plt.plot(year,value)
        plt.title(u'{}'.format(sheet.row_values(0)[i+1]))
        #plt.show()
        plt.savefig(picfile+u'{}'.format(sheet.row_values(0)[i+1]))
        plt.cla()

def Genre():
    for i in range(1, 13, 2):
        value = []
        dic = {}
        for j in range(clen):
            if sheet.row_values(i+1)[j+1] == 0:
                value.append(0)
            else:
                value.append(sheet.row_values(i)[j+1] / sheet.row_values(i+1)[j+1])
        for k in range(len(value)):
            dic[genre[k]] = value[k]
        l = sorted(dic.items(), key=lambda x: x[1], reverse=True)
        x,y = [],[]
        for n in range(len(l)):
            x.append(l[n][0])
            y.append(l[n][1])
        plt.rcParams['figure.figsize'] = (10,8)
        plt.plot(x,y)
        plt.xticks(rotation=90)
        plt.title(u'{}'.format(sheet.col_values(0)[i]))
        #plt.show()
        plt.savefig(picfile + u'{}'.format(int(sheet.col_values(0)[i])))
        plt.cla()

def Genre_all():
    dic = {}
    for i in range(1, 13, 2):
        value = []
        for j in range(clen):
            if sheet.row_values(i + 1)[j + 1] == 0:
                value.append(0)
            else:
                value.append(sheet.row_values(i)[j + 1] / sheet.row_values(i + 1)[j + 1])
        for k in range(len(value)):
            if genre[k] not in dic:
                dic[genre[k]] = value[k]
            else:
                dic[genre[k]] += value[k]
    l = sorted(dic.items(), key=lambda x: x[1], reverse=True)
    x, y = [], []
    for n in range(len(l)):
        x.append(l[n][0])
        y.append(l[n][1])
    plt.rcParams['figure.figsize'] = (10, 8)
    plt.plot(x, y)
    plt.xticks(rotation=90)
    # plt.title(u'{}'.format(sheet.col_values(0)[i]))
    plt.savefig('all')
    plt.show()
    # plt.cla()

# Year()
# Genre()
Genre_all()