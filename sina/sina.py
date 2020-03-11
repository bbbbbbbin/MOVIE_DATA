# -*- coding: utf-8 -*-
import re
import json
from urllib import quote_plus
import requests
import base64
import rsa
import binascii
import time
import logging
import urllib2
import xlrd
from xlutils.copy import copy
import datetime
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import requests
import random
from lxml import etree
from selenium import webdriver

def get_search_number(keyword, before7, release):
    # url = ''.format('http://s.weibo.com/wb/{0}',urllib.quote('爸爸去哪儿&xsort=time&timescope=custom:2014-01-24:2014-01-31&Refer=g'))
    parameters = '{0}'.format(keyword)
    url = 'http://weibo.cn/search/mblog/?keyword=' + urllib2.quote(parameters) + '&rl=1&starttime={0}&endtime={1}&sort=time'.format(before7, release)
    # url = 'http://weibo.cn/search/mblog/?keyword={0}&rl=1&starttime={1}&endtime={2}&sort=time'.format(keyword, before7, release)
    print url
    cookies = {
        'Cookie': "WEIBOCN_WM=3333_2001; _T_WM=c881970296f203298da2a0ee672499e6; ALF=1535118726; SCF=AllczJZv5Zpsu0RrXisqYx_ZFu5pu0KNuqxX0OavXkBbkJZeKNx-eGwQvJiiBqJY88HrUtNaW2pCRlQmymK-o7E.; SUB=_2A252XPDfDeRhGeNK61YX9ibFwjSIHXVVvpCXrDV6PUJbktANLVfMkW1NSW5fIDkDSgf_ybXDnfukk0mcfLQRtYTu; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WW9Im.Nj8aAJms3gZao5ZGe5JpX5K-hUgL.Fo-XehBcSon41Kn2dJLoI0YLxKqL12BLBKzLxK-LB-BLBKeLxK-LBo5L1K2LxK-LBo.LBoBLxK-L1KeL1hnLxK-L1KeL1hnLxK-L1KeL1hnt; SUHB=0WGchowu2TcdDX; SSOLoginState=1532526735"}
    headers = {
        "User-Agent": 'Mozilla/5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 66.0.3359.181Safari / 537.36'}
    try:
        html = requests.get(url,headers=headers,cookies=cookies)
        time.sleep(random.random() * 5)
        data = html.content
        et = etree.HTML(data)
        info = et.xpath("//div")[5]
        matched = info.xpath("./span/text()")
        n = re.compile(r'(\w*[0-9]+)\w*')
        nums = n.findall(matched[0])[0]
        nums = eval(nums)
        if nums == 0:
            return -1
        else:
            return nums
    except:
        # print keyword, before7, release, '=> error'
        return -1

def get_weibo_index(excel_path):
    rbook = xlrd.open_workbook(excel_path)
    rsheet = rbook.sheet_by_index(0)
    wbook = copy(rbook)
    wsheet = wbook.get_sheet(0)
    keywords = rsheet.col_values(1)  # F列
    releases = rsheet.col_values(5)  # G列
    weibo_search_column = 14
    # release_date = datetime.datetime.strptime(str,'%Y-%m-%d')
    # before7 = release_date - datetime.timedelta(days=7)
    # before7.strftime('%Y-%m-%d')

    for i in range(len(keywords)):
        keyword = keywords[i]
        release = releases[i]
        if release == "":
            continue
        old = rsheet.cell_value(i,weibo_search_column)
        if old != -1:
            continue
        else:
            release_date = datetime.datetime.strptime(str(release).strip(), '%Y-%m-%d')
            before1_date = release_date - datetime.timedelta(days=1)
            before1 = before1_date.strftime('%Y%m%d')
            before7_date = release_date - datetime.timedelta(days=90)
            before7 = before7_date.strftime('%Y%m%d')
            print keyword, before7, before1, release
            num = get_search_number(keyword, before7, before1)
            print num
            wsheet.write(i, weibo_search_column, num)
        # if old != -1:
        #     if old >= num:
        #         print old
        #         wsheet.write(i, weibo_search_column, old)
        #     else:
        #         print num
        #         wsheet.write(i, weibo_search_column, num)
        # else:
        #     print num
        #     wsheet.write(i, weibo_search_column, num)
    wbook.save(excel_path)

# chromePath = r'../chromedriver.exe'
# wd = webdriver.Chrome(executable_path= chromePath) #构建浏览器
# loginUrl = 'https://www.weibo.com'
# wd.get(loginUrl) #进入登陆界面
# time.sleep(10)
# # wd.find_element_by_css_selector("#loginname").send_keys("18805053417")
# # wd.find_element_by_css_selector(".info_list.password input[node-type='password']").send_keys("Hecb0071997")
# # wd.find_element_by_css_selector(".info_list.login_btn a[node-type='submitBtn']").click()
# wd.find_element_by_xpath('//*[@id="loginname"]').send_keys('18805053417') #输入用户名
# wd.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[2]/div/input').send_keys('Hecb0071997') #输入密码
# wd.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[6]/a').click() #点击登陆
# wd.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[3]/div/input').send_keys(raw_input("输入验证码： "))
# wd.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[3]/div[6]/a').click()#再次点击登陆
# time.sleep(5)
# req = requests.Session() #构建Session
# cookies = wd.get_cookies() #导出cookie
# for cookie in cookies:
#     req.cookies.set(cookie['name'],cookie['value']) #转换cookies
# test = req.get('http://weibo.cn/search/mblog/?keyword=%E5%8F%98%E5%BD%A2%E9%87%91%E5%88%9A4&rl=1&starttime=2014-06-20&endtime=2014-06-26&sort=time')
# data = test.content
# print data

#for j in range(0,5):
for i in range(0, 8):
    get_weibo_index('data/{}.xls'.format(2010+i))
    #time.sleep(60)