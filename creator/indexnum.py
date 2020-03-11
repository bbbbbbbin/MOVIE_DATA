# -*- coding:utf -8-*-

import xlwt
import json
import simplejson
import sys
import numpy as np
reload(sys)
sys.setdefaultencoding('utf8')

jsonfile = r'json/'

f = open(jsonfile+r'creator.json')
creator = simplejson.load(f)
f1 = open(jsonfile+r'w.json')
w = simplejson.load(f1)
f.close()
f1.close()

w_year = w['year']  #6
w_act = w['act']    #4
w_year = np.array(w_year)
w_act = np.array(w_act)
w_year = w_year.reshape(1,6)
w_act = w_act.reshape(4,1)
#print w_year,w_act

def cal(value):
    a = np.dot(w_year,value)
    a = np.dot(a,w_act)
    return a[0][0]

new = {}
for crea,value in creator.items():
    value = np.array(value)
    new[crea] = cal(value)

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('act')
k = -1
for i,j in new.items():
    k = k+1
    sheet.write(k, 0, i)
    sheet.write(k, 1, j)
book.save(jsonfile+r'act.xls')


