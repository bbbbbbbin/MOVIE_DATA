# -*- coding:utf -8-*-

import xlrd
import xlwt
import sys
reload(sys)
sys.setdefaultencoding('utf8')

xlsfile = r'data/'

def genre_list():
    g_l = []
    for i in range(0,6):
        book = xlrd.open_workbook(xlsfile + '{}.xls'.format(2010+i))
        sheet = book.sheet_by_index(0)
        genres = sheet.col_values(13)
        for genre in genres:
            gs = genre.split('/')
            for g in gs:
                if g not in g_l:
                    g_l.append(g)
    return g_l

def write_dic():
    genre_box = {}
    for k in range(0,6):
        book = xlrd.open_workbook(xlsfile + '{}.xls'.format(2010+k))
        sheet = book.sheet_by_index(0)
        genres = sheet.col_values(13)
        box = sheet.col_values(2)
        i = 0
        for genre in genres:
            gs = genre.split('/')
            for g in gs:
                if g+'_{}'.format(2010+k) not in genre_box:
                    genre_box[g+'_{}'.format(2010+k)] = 0 + box[i]
                    genre_box[g+'_{}_n'.format(2010+k)] = 1
                else:
                    genre_box[g + '_{}'.format(2010+k)] += box[i]
                    genre_box[g + '_{}_n'.format(2010+k)] += 1
            i += 1
    return genre_box

def writetitle(gl,nw):
    i = 1
    for l in gl:
        nw.write(0,i,l)
        i+=1

gl = genre_list()
gb = write_dic()
nb = xlwt.Workbook(encoding='utf-8', style_compression=0)
nw = nb.add_sheet('all')
writetitle(gl,nw)
j = 1
for i in range(0,6):
    nw.write(j,0,2010+i)
    k = 1
    for l in gl:
        if l+'_{}'.format(2010+i) in gb:
            nw.write(j,k,gb[l+'_{}'.format(2010+i)])
            nw.write(j+1, k, gb[l + '_{}_n'.format(2010 + i)])
        else:
            nw.write(j,k,0)
            nw.write(j+1,k,0)
        k += 1
    j += 2
nb.save(xlsfile+'produce.xls')
