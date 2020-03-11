# -*- coding:utf -8-*-

import xlrd
import json
import sys
reload(sys)
sys.setdefaultencoding('utf8')

xlsfile = r'../data/exact/'
jsonfile1 = r'json/'

creator = {}
box_all = []
boxall = 0
for k in range(0,6):
    book = xlrd.open_workbook(xlsfile+'{}.xls'.format(2010+k))
    sheet = book.sheet_by_index(0)
    idxs = sheet.col_values(0)
    box_year = 0
    for i in range(len(idxs)):
        boxoffice = sheet.col_values(2)[i]
        box_year += boxoffice
        for j in range(7,13):
            act = sheet.col_values(j)[i]
            if act != "":
                if act not in creator:
                    creator[act] = [[0, 0, 0, 0],
                                    [0, 0, 0, 0],
                                    [0, 0, 0, 0],
                                    [0, 0, 0, 0],
                                    [0, 0, 0, 0],
                                    [0, 0, 0, 0]]
                if j == 7:
                    creator[act][k][0] += boxoffice
                elif j == 8:
                    creator[act][k][1] += boxoffice
                elif j == 9 or j == 10:
                    creator[act][k][2] += boxoffice
                elif j == 11 or j == 12:
                    creator[act][k][3] += boxoffice
    boxall += box_year
    box_all.append(box_year)

f = open(jsonfile1+'creator.json','w')
f.write(json.dumps(creator, indent=4, ensure_ascii=False))
f.close()

w = {}
for idx, box in enumerate(box_all):
    box_all[idx] = box / boxall
w['year'] = box_all
w['act'] = [0.417, 0.167, 0.333, 0.083]
f1 = open(jsonfile1+'w.json','w')
f1.write(json.dumps(w, indent=4, ensure_ascii=False))
f1.close()
