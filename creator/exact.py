# -*- coding:utf-8 -*-

import xlrd
import json
import re
import xlwt
import simplejson
from xlutils.copy import copy

def Todic(file):
    dic = {}
    rd = xlrd.open_workbook(file)
    wd = rd.sheet_by_index(0)
    name = wd.col_values(0)
    value = wd.col_values(1)
    for i in range(len(value)):
        dic[name[i]] = value[i]
    return dic

def Changcrea(file, dic):
    rd = xlrd.open_workbook(file)
    wd = rd.sheet_by_index(0)
    newrd = copy(rd)
    newwd = newrd.get_sheet(0)
    id = wd.col_values(0)
    for i in range(len(id)):
        for j in range(7,13):
            name = wd.col_values(j)[i]
            if name not in dic:
                newwd.write(i, j, 0)
            else:
                newwd.write(i, j, dic[name])
    newrd.save(file)

dic = Todic('json/act.xls')
Changcrea('data/2016.xls', dic)
Changcrea('data/2017.xls', dic)
