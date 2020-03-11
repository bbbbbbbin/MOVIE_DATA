# -*- coding:utf-8 -*-

from PIL import Image,ImageFilter,ImageEnhance
import numpy as np
import xlrd
import xlwt
import matplotlib.image as mpimg
import pytesseract
import os
import time
import datetime
from aip import AipOcr
import sys
reload(sys)
sys.setdefaultencoding('utf8')


config = {
    'appId':'11592499',
    'apiKey':'euAlypGmSOG0SVdcUhSbygPS',
    'secretKey':'0XSY30XtnFoz62ssdpfTUaBridzvEeGd'
}
client = AipOcr(**config)
def picp(file):
    pic = Image.open(file)
    pic = pic.resize((75,18))
    pic = pic.convert('RGBA')
    r,g,b,a = pic.split()
    t = 100
    a = np.array(a)
    r = np.array(r)
    g = np.array(g)
    b = np.array(b)
    a = np.zeros(a.shape)
    for i in range(a.shape[0]):
        for j in range(a.shape[1]):
            if r[i][j] <= t and g[i][j] <= t and b[i][j] <= t:
                a[i][j] = 255
    a = Image.fromarray(a)
    r = Image.fromarray(r)
    g = Image.fromarray(g)
    b = Image.fromarray(b)
    a = a.convert('L')
    r = r.convert('L')
    g = g.convert('L')
    b = b.convert('L')
    pic = Image.merge('RGBA',(r,g,b,a))
    pic = pic.filter(ImageFilter.SHARPEN)
    enhancer = ImageEnhance.Contrast(pic)
    pic = enhancer.enhance(1.8)
    code = pytesseract.image_to_string(pic,lang='chi_sim')
    pic.save(file.split('.')[0]+'.png')
    return code

book = xlrd.open_workbook('data/2016.xls')
sheet = book.sheet_by_index(0)
idxs = sheet.col_values(0)
names = sheet.col_values(1)
dates = sheet.col_values(5)
nb = xlwt.Workbook(encoding='utf-8', style_compression=0)
nw = nb.add_sheet('all')
for i in range(len(idxs)):
    path = '2016/{}/'.format(idxs[i])
    folder = os.path.exists(path)
    date = dates[i]
    nw.write(i, 0, idxs[i])
    nw.write(i, 1, names[i])
    if not folder:
        nw.write(i, 2, 0)
    else:
        j = 0
        year_r = int(date.split('-')[0])
        month_r = int(date.split('-')[1])
        day_r = int(date.split('-')[2])
        day1 = datetime.date(year_r,month_r,day_r)
        print day1
        for filename in os.listdir(path):
            filename_r = filename.split('.')[0]
            year = int(filename_r.split('-')[0])
            month = int(filename_r.split('-')[1])
            day = int(filename_r.split('-')[2])-1
            day2 = datetime.date(year, month, day)
            daydel = day2-day1
            if  -30 <= int(daydel.days) < 0:
                pic = open(path+filename_r+'.jpg','rb')
                pic = pic.read()
                try:
                    codes = client.basicGeneral(pic)
                    if 'words_result' in codes:
                        try:
                            code = codes['words_result'][0]['words']
                            nw.write(i, j + 2, code)
                            j += 1
                        except:
                            code = ''
                    print code
                except:
                    nb.save('data/all.xls')

nb.save('data/all.xls')
