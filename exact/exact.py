# -*- coding:utf-8 -*-

import xlrd
import json
import re
import xlwt
import simplejson
from xlutils.copy import copy

xlsfile = r'../data/douban/'
#jsonfile1 = r'../data/movie/'
jsonfile2 = r'../data/movie2/'
save_file = r'../data/exact/'


def min(path):
    rb = xlrd.open_workbook(path)
    rs = rb.sheet_by_index(0)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    durations = rs.col_values(4)
    k = -1
    for duration in durations:
        k = k+1
        p = re.compile(r'(\w*[0-9]+)\w*')
        mi = p.findall(duration.encode('utf-8'))
        ws.write(k, 4, eval(mi[0]))
    wb.save(path)

def BoxstrToNum(boxoffice):
    n = re.compile(r'(\w*[0-9]+)\w*')
    b = re.compile(r'[^>]+亿')
    t = re.compile(r'[^>]+万')
    nums = n.findall(boxoffice)
    if b.match(boxoffice):
        number = ''
        for num in nums:
            number = number+num
        if len(nums) == 1:
            number = number + '0000'
        elif len(nums) == 2:
            if len(nums[1]) == 1:
                number = number + '000'
            elif len(nums[1]) == 2:
                number = number + '00'
            elif len(nums[1]) == 3:
                number = number +'0'
            else:
                number = '0'
        else:
            number = '0'
        number = eval(number)
        return number
    elif t.match(boxoffice):
        if len(nums) == 2:
            number = eval(nums[0]+'.'+nums[1])
        elif len(nums) == 1:
            number = eval(nums[0])
        else:
            number = 0
        return number

def JsonToExcel(year):
    # with open(jsonfile1+'movies{}.json'.format(2017),'r') as f:
    #    dic1 = json.load(f)
    with open(jsonfile2+'movies_date{}.json'.format(year),'r') as f:
        dic2 = simplejson.load(f)
    book = xlrd.open_workbook(xlsfile+'boxoffice_{}.xls'.format(year))
    book2 = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = book.sheet_by_index(0)
    sheet2 = book2.add_sheet('{}'.format(year),cell_overwrite_ok=True)
    movies = sheet.col_values(0)
    boxoffice = sheet.col_values(1)
    idxs = sheet.col_values(4)
    p = re.compile(r'(\w*[0-9]+)\w*')
    idxs = [ p.findall(str(idx))[0] for idx in idxs]
    for i in range(len(idxs)):
        sheet2.write(i, 0, idxs[i])
        sheet2.write(i, 1, movies[i])
        boxnum = BoxstrToNum(boxoffice[i].encode('utf-8'))
        sheet2.write(i, 2, boxnum)
        wish = dic2[idxs[i]]['wishes']
        sheet2.write(i, 3, wish)
        duration = dic2[idxs[i]]['duration']
        sheet2.write(i, 4, duration)
        release = dic2[idxs[i]]['release']
        sheet2.write(i, 5, release)
        genres = dic2[idxs[i]]['genre']
        if genres != []:
            ge = ""
            for genre in genres:
                ge = ge + genre + '/'
            sheet2.write(i, 6, ge)
        else:
            sheet2.write(i, 6, None)
        director = dic2[idxs[i]]['directors']
        if director != []:
            sheet2.write(i, 7, director[0])
        else:
            sheet2.write(i, 7, None)
        write = dic2[idxs[i]]['writer']
        if write != []:
            sheet2.write(i, 8, write[0])
        else:
            sheet2.write(i, 8, None)
        casts = dic2[idxs[i]]['actors']
        if len(casts) >= 4:
            sheet2.write(i, 9, casts[0])
            sheet2.write(i, 10, casts[1])
            sheet2.write(i, 11, casts[2])
            sheet2.write(i, 12, casts[3])
        elif len(casts) == 3:
            sheet2.write(i, 9, casts[0])
            sheet2.write(i, 10, casts[1])
            sheet2.write(i, 11, casts[2])
        elif len(casts) == 2:
            sheet2.write(i, 9, casts[0])
            sheet2.write(i, 10, casts[1])
        elif len(casts) == 1:
            sheet2.write(i, 9, casts[0])
        else:
            sheet2.write(i, 9, None)

    book2.save(save_file+r'{}.xls'.format(year))

#for i in range(1,8):
    #JsonToExcel(2010+i)

#print BoxstrToNum(u'1.24亿'.encode('utf-8'))
    #min(r'../data/exact/{}.xls'.format(2010+i))