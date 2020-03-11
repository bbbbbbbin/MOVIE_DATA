# -*- coding:utf-8 -*-
# 此程序用于读取数据库电影名与日期并且爬取该电影的百度指数，然后保存到数据库中
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
# import ReadXml
# from ReadXml import getFirstLvValue
import time
from PIL import Image
import pytesseract
import re
import os
from selenium.common.exceptions import StaleElementReferenceException
import random
import xlrd
from xlutils.copy import copy
import pickle
from PIL import ImageOps
from aip import AipOcr
# 全局常量
# 截取图片保存路径
path = os.getcwd()
# 月份-日字典
Monthdict = {'01': 31, '02': 28, '03': 31, '04': 30, '05': 31, '06': 30, '07': 31, '08': 31, '09': 30, '10': 31,
             '11': 30, '12': 31, '1': 31, '2': 28, '3': 31, '4': 30, '5': 31, '6': 30, '7': 31, '8': 31, '9': 30}

# 账号地点，保存多个账号用于随机选取
AccountList = [['18805053417', 'HECB0071997']]

# xml路径
XmlPath = "data/boxoffice.xls"

# 全局变量
inputid = ''
name = ''

config = {
    'appId':'11592499',
    'apiKey':'euAlypGmSOG0SVdcUhSbygPS',
    'secretKey':'0XSY30XtnFoz62ssdpfTUaBridzvEeGd'
}
client = AipOcr(**config)
# ------------------------------------Spider流程#------------------------------------#

# 初始化数据库并且创建文件本地保存目录
# def init_sys():
#     # try:
#         #读取Xml并初始化SQL
#         ReadXml.init_path(XmlPath)
#         SQLTools.InitSql(getFirstLvValue('host'),getFirstLvValue('user'),getFirstLvValue('passwd'),getFirstLvValue('db'),getFirstLvValue('charset'))
#         #创建文件本地保存目录
#         if(os.path.exists(path+"/raw")) is False:
#             os.mkdir(path+"/raw")
#         if (os.path.exists(path + "/crop")) is False:
#             os.mkdir(path + "/crop")
#         if (os.path.exists(path + "/zoom")) is False:
#             os.mkdir(path + "/zoom")
#         return True
# except Exception,e :
#     print e.message
#     return False

def initTable(threshold=140):  # 二值化函数
    table = []
    for i in range(256):
        if i < threshold:
            table.append(0)
        else:
            table.append(1)

    return table


# 从数据库读取任务,返回Request list
def load_req(i, keywords, releases):
    global name
    # 获取任务
    #  G列
    # (keyword,time)
    if keywords != "":
        name = keywords[i]
        print keywords[i]
        day = releases[i].split('-')[2]
        month = releases[i].split('-')[1]
        year = releases[i].split('-')[0]
        print "正在获取", name.encode("utf-8"), "的百度指数"
        return [name, year, month, day]
    else:
        return False


# 初始化Spiderk 模拟登录 提供browser工具类
def init_spider():
    try:
        url = "http://index.baidu.com/"  # 百度指数网站
        chromePath = r'../chromedriver.exe'
        browser = webdriver.Chrome(executable_path=chromePath)
        browser.get(url)
        # 点击网页的登录按钮
        browser.find_element_by_xpath("//span[@class='username-text']").click()
        time.sleep(3)
        # 传入账号密码
        list = random.choice(AccountList)
        try:
            browser.find_element_by_id("TANGRAM__PSP_4__password").send_keys(list[1])
            browser.find_element_by_id("TANGRAM__PSP_4__userName").send_keys(list[0].encode("utf-8"))
            browser.find_element_by_id("TANGRAM__PSP_4__submit").click()
            a = raw_input("")
        except:
            browser.find_element_by_id("TANGRAM__PSP_4__password").send_keys(list[1])
            browser.find_element_by_id("TANGRAM__PSP_4__userName").send_keys(list[0])
            browser.find_element_by_id("TANGRAM__PSP_4__submit").click()
        time.sleep(3)
        pickle.dump(browser.get_cookies(),open("cookies.pkl", 'wb'))
        return browser
    except:
        return False


# 执行Spider 返回数据结果
def exec_spider(request):
    # request(name,year,month,day)
    global browser
    try:
        name = request[0]
        year = request[1]
        month = request[2]
        day = request[3]
        # 清空网页输入框
        try:
            browser.find_element_by_id("schword").clear()
            # 写入需要搜索的百度指数
            browser.find_element_by_id("schword").send_keys(name)
        except:
            browser.find_elements_by_id("search-input-word").clear()
            # 写入需要搜索的百度指数
            browser.find_element_by_id("search-input-word").send_keys(name)
        # 点击搜索
        try:
            browser.find_element_by_id("searchWords").click()
        except:
            browser.find_element_by_id("schsubmit").click()
        time.sleep(2)

        fyear, fmonth, ayear, amonth = CalculateDate(year, month)
        # 点击网页上的开始日期
        if str(fyear) == "2010":
            return False
        browser.maximize_window()
        browser.find_elements_by_xpath("//div[@class='box-toolbar']/a")[6].click()
        browser.find_elements_by_xpath("//span[@class='selectA yearA']")[0].click()
        browser.find_element_by_xpath(
            "//span[@class='selectA yearA slided']//div//a[@href='#" + str(fyear) + "']").click()
        browser.find_elements_by_xpath("//span[@class='selectA monthA']")[0].click()
        browser.find_element_by_xpath(
            "//span[@class='selectA monthA slided']//ul//li//a[@href='#" + str(fmonth) + "']").click()
        # 选择网页上的截止日期
        browser.find_elements_by_xpath("//span[@class='selectA yearA']")[1].click()
        browser.find_element_by_xpath(
            "//span[@class='selectA yearA slided']//div//a[@href='#" + str(ayear) + "']").click()
        browser.find_elements_by_xpath("//span[@class='selectA monthA']")[1].click()
        browser.find_element_by_xpath(
            "//span[@class='selectA monthA slided']//ul//li//a[@href='#" + str(amonth) + "']").click()
        browser.find_element_by_xpath("//input[@value='确定']").click()
        time.sleep(2)

        # 闰年处理
        if int(year) == 2012 or int(year) == 2016:
            Monthdict['02'] = 29

        return CollectIndex(browser, fyear, fmonth, day, name)
    except IndexError, e:
        print e
        if Anti_Exist(browser) is True:
            browser.close()
            time.sleep(100)
            browser = init_spider()
            return exec_spider(request)
        else:
            return False
    except StaleElementReferenceException, e2:
        browser.close()
        time.sleep(100)
        browser = init_spider()
        return exec_spider(request)


# 对抗反爬机制
def Anti_Exist(browser):
    try:
        browser.find_element_by_xpath("//img[@src='/static/imgs/deny.png']")
        return True
    except:
        return False


# ----------------------------------Spider运行过程中所需的方法----------------------------------#

# 计算需要选择的日期——电影上映前后一个月
def CalculateDate(year, month):
    if year == '2010':
        fyear = 2011
        fmonth = '01'
    else:
        fyear = year
        if int(month) == 1:
            fmonth = '12'
            fyear = str(int(year) - 1)
        else:
            fmonth = str(int(month) - 1)
    if len(fmonth) < 2:
        fmonth = '0' + fmonth
    if year == '2010':
        ayear = 2011
        amonth = '03'
    else:
        ayear = year
        if int(month) + 1 == 13:
            amonth = '01'
            ayear = str(int(year) + 1)
        else:
            amonth = str(int(month) + 1)
    if len(amonth) < 2:
        amonth = '0' + amonth
    return fyear, fmonth, ayear, amonth


def CollectIndex(browser, fyear, fmonth, day, name):
    # 初始化输出String
    OutputString = '['
    x_0 = 1
    y_0 = 1
    # 根据起始具体日子计算鼠标的初始位置
    # 一日=13.51 例如,上映日期为7.20日 则x起始坐标为1+13.41*19
    if str(fyear) != '2011':
        ran = Monthdict[fmonth] + int(day) - 32
        if ran < 0:
            ran = 0.5
        x_0 = x_0 + 13.51 * ran
    else:
        day = 1
    xoyelement = browser.find_elements_by_css_selector("#trend rect")[2]
    ActionChains(browser).move_to_element_with_offset(xoyelement, x_0, y_0).perform()
    for i in range(61):
        # 计算当前得到指数的时间
        if int(fmonth) < 10:
            fmonth = '0' + str(int(fmonth))
        if int(day) >= Monthdict[str(fmonth)] + 1:
            day = 1
            fmonth = int(fmonth) + 1
            if fmonth == 13:
                fyear = int(fyear) + 1
                fmonth = 1
        day = int(day) + 1
        time.sleep(1)
        # 获取Code
        code = GetTheCode(browser, fyear, fmonth, day, name, path, xoyelement, x_0, y_0)
        # ViewBox不出现的循环
        cot = 0
        jud = True
        # print code
        while (code == None):
            cot += 1
            code = GetTheCode(browser, fyear, fmonth, day, name, path, xoyelement, x_0, y_0)
            if cot >= 6:
                jud = False
                break
        if jud:
            anwserCode = code.group()
        else:
            anwserCode = str(-1)
            if int(day) < 10:
                day = '0' + str(int(day))
            if int(fmonth) < 10:
                fmonth = '0' + str(int(fmonth))
        OutputString += str(fyear) + '-' + str(fmonth) + '-' + str(day) + ':' + str(anwserCode) + ','
        x_0 = x_0 + 13.51
        print anwserCode
    OutputString += ']'
    print OutputString
    return OutputString.decode('utf-8')


def GetTheCode(browser, fyear, fmonth, day, name, path, xoyelement, x_0, y_0):
    ActionChains(browser).move_to_element_with_offset(xoyelement, x_0, y_0).perform()
    # 鼠标重复操作直到ViewBox出现
    cot = 0
    while (ExistBox(browser) == False):
        cot += 1
        ActionChains(browser).move_to_element_with_offset(xoyelement, x_0, y_0).perform()
        if ExistBox(browser) == True:
            break
        if cot == 6:
            return None

    imgelement = browser.find_element_by_xpath('//div[@id="viewbox"]')
    locations = imgelement.location
    printString = str(fyear) + "-" + str(fmonth) + "-" + str(day)
    # 找到图片位置
    l = len(name)
    if l > 8:
        l = 8
    rangle = (int(int(locations['x'])) + l * 10 + 42, int(int(locations['y'])) + 33,
              int(int(locations['x'])) + l * 10 + 42 + 75,
              int(int(locations['y'])) + 51)
    # 保存截图
    browser.save_screenshot(str(path) + "/raw/" + printString + ".png")
    img = Image.open(str(path) + "/raw/" + printString + ".png")
    if locations['x'] != 0.0:
        # 按Rangle截取图片
        jpg = img.crop(rangle)
        imgpath = str(path) + "/crop/" + printString + ".jpg"
        r, g, b, a = jpg.split()
        jpg = Image.merge("RGB", (r, g, b))
        jpg.save(imgpath)
        jpgzoom = Image.open(str(imgpath))
        # 放大图片
        out = jpgzoom.convert('L')
        out = out.point(initTable(), '1')
        out = out.convert('L')
        out = ImageOps.invert(out)
        out = out.convert('1')
        out = out.convert('L')
        (x, y) = jpgzoom.size
        x_s = x * 10
        y_s = y * 10
        out = out.resize((x_s, y_s), Image.ANTIALIAS)
        out.save(path + "/zoom/" + printString + ".jpg", 'jpeg', quality=95)
        #image = Image.open(path + "/zoom/" + printString + ".jpg")
        image = open(path + "/zoom/" + printString + ".jpg", 'rb')
        bu = image.read()
        # 识别图片
        codes = client.basicGeneral(bu)
        time.sleep(0.5)
        if 'words_result' in codes:
            try:
                code =  codes['words_result'][0]['words']
            except:
                code = ''
        #code = pytesseract.image_to_string(image, nice=10)
        regex = "\d+"
        pattern = re.compile(regex)
        dealcode = code.replace("S", '5').replace(" ", "").replace(",", "").replace("E", "8").replace(".", ""). \
            replace("'", "").replace(u"‘", "").replace("B", "8").replace("\"", "").replace("I", "1").replace(
            "i", "").replace("-", ""). \
            replace("$", "8").replace(u"’", "").strip()
        match = pattern.search(dealcode)
        return match
    else:
        return None


# 判断ViewBox是否存在
def ExistBox(browser):
    try:
        browser.find_element_by_xpath('//div[@id="viewbox"]')
        return True
    except:
        return False


# -------------------------------------主函数代码-----------------------------------------#

if __name__ == '__main__':
    global browser
    # status记录初始化状态
    # status=init_sys()
    # if status is False:
    #    exit(1)
    # status记录工具类
    url = "http://index.baidu.com/"  # 百度指数网站
    chromePath = r'../chromedriver.exe'
    browser = webdriver.Chrome(executable_path=chromePath)
    cookies = pickle.load(open("cookies.pkl", "rb"))
    browser.get(url)
    for cookie in cookies:
        browser.add_cookie(cookie)
    status = browser
    b = raw_input()
    # status = init_spider()
    rbook = xlrd.open_workbook(XmlPath)
    rsheet = rbook.sheet_by_index(0)
    keywords = rsheet.col_values(0)  # F列
    releases = rsheet.col_values(5)
    wbook = copy(rbook)
    wsheet = wbook.get_sheet(0)
    ticot = 0
    if status is False:
        exit(3)
    else:
        browser = status
    while True:
        ticot += 1
        # SQLTools.Renew()
        # status记录Request
        status = load_req(ticot, keywords, releases)
        if status is False:
            exit(2)
        else:
            request = status
            # request=[name,year,month,day,id]
        # status记录结果
        status = exec_spider(request)
        if status is False:
            print name, 'Error'
            wsheet.write(ticot, 8, -1)
            continue
        else:
            wsheet.write(ticot, 8, status)
        print "将结果保存到数据库", ticot
        # 保存到数据库中
        # SQLTools.SaveResultToDB(resultString,request[4])
        # SQLTools.AlterStatus("update baidu_index set status=1 where input_id=" + str(request[4]) + ";")
        # 获取下一条
        print "休息"
        time.sleep(10)
        print "获取下一条数据"
