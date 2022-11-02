# -*- coding = utf-8 -*-
from shutil import copyfile
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.options import Options  # 手机模式
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import json
import time
import os
import xlrd
from PIL import Image, ImageFont, ImageDraw
from openpyxl.drawing.image import Image as Image2

import win32com.client as win32

# 设置手机型号，这设置为iPhone 6
mobile_emulation = {"deviceName": "iPhone 6"}

options = Options()
options.add_experimental_option("mobileEmulation", mobile_emulation)

service = Service(r"D:\Program Files\Chrome\Application\chromedriver.exe")
brguge = webdriver.Chrome(service=service, options=options)


def shotScreen(obj, rown, celln):
    try:
        brguge.get('https://yqpt.xa.gov.cn/nrt/inquire.html')
        wait = WebDriverWait(brguge, 10)
        name = brguge.find_element('id', 'personName')

        name.send_keys(obj['nameInput'])

        personIdcard = brguge.find_element('id','personIdcard')
        personIdcard.send_keys(obj['cardInput'])

        submitBtn = brguge.find_element('id','submitBtn')
        submitBtn.send_keys(Keys.ENTER)

        drag_btn = brguge.find_element(By.CLASS_NAME,'drag-btn')

        brguge.execute_script('localStorage.setItem("inquire-name", "'+obj['nameInput']+'")')
        brguge.execute_script('localStorage.setItem("inquire-card", "' + obj['cardInput'] + '")')
        brguge.execute_script('window.location.href = "https://yqpt.xa.gov.cn/nrt/resultQuery.html"')


        width = brguge.execute_script("return document.documentElement.scrollWidth")
        height = brguge.execute_script("return document.documentElement.scrollHeight")
        # 将窗口设置为页面滚动宽高
        brguge.set_window_size(width, height)


        filepath = foldername + "\\" + obj['nameInput'] + '.png'
        brguge.save_screenshot(filepath)

        makeWater(filepath, filepath, obj['nameInput'] + ','+ obj['relationship'])

        # insert the picture
        wb = load_workbook(excel, data_only=True)
        sheet = wb['学校明细']
        img = Image2(filepath)

        _from = AnchorMarker(celln, 30, rown, 30)  # 创建锚标记对象,设置图片所占的row
        to = AnchorMarker(celln + 1, -30, rown + 1, -30)  # 创建锚标记对象,设置图片所占的row 从而确认了图片位置
        img.anchor = TwoCellAnchor('twoCell', _from, to)

        #sheet.add_image(img, celln + str(rown+1))
        sheet.add_image(img)
        wb.save(excel)
        wb.close()

    except Exception as e:
        print(e)
        #brguge.close()





# 加水印
# destPath 保存位置
# img 原图路径
# txtWater 水印文字
# waterPic 水印图片
def makeWater(destPath, img, txtWater):
    im = Image.open(img)

    # 加图片水印
    #mark = Image.open(waterPic)
    layer = Image.new('RGBA', im.size, (0, 0, 0, 0))
    #layer.paste(im, (0, im.size[1] - im.size[1] - 10))  # 水印图位置（右下角）
    out = Image.composite(layer, im, layer)

    # 加文字水印
    # 设置所使用的字体
    font = ImageFont.truetype("C:\Windows\Fonts\STXIHEI.TTF", 32)
    draw = ImageDraw.Draw(out)
    # 坐标,水印文字,文字颜色,字体
    draw.text((40, im.size[1] - 100), txtWater, (123, 123, 132), font=font)  # 文字水印位置（左下角）
    draw = ImageDraw.Draw(out)

    # 保存图片
    out.save(destPath)

def exchange(files):

    for file_name in files:
        if file_name.rsplit('.',1)[-1]=='xls':
            fname = os.path.join(path,file_name)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)
            #在原来的位置创建出：原名+'.xlsx'文件
            wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
            wb.Close()                               #FileFormat = 56 is for .xls extension
            excel.Application.Quit()
            os.remove(fname)


with open("namelist.json", encoding="utf-8") as f:
    objs = json.load(f)

folder_date = time.strftime('%Y-%m-%d', time.localtime(time.time()))

path = r"D:\\"
foldername = path + "\\" + folder_date
isExists = os.path.exists(foldername)
if not isExists:
    print("目录不存在")
    os.makedirs(foldername)
else:
    print("目录" + ":  " + foldername + "  已存在")

source = 'D:\\Hesuan.xls'
excel = 'D:\\Hesuan2.xls'
if (os.path.exists(excel)):
    os.remove(excel)
copyfile(source, excel)

data = xlrd.open_workbook(excel)
table =data.sheets()[0]

exchange([excel])
excel = 'D:\\Hesuan2.xlsx'

def getRow(name):
    for rown in range(table.nrows):
        cellName = table.cell_value(rown, 1)
        if (name == cellName):
            return rown


az = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

for i in range(len(objs)):

    row = getRow(objs[i]['primary']['nameInput'])
    # 生成小孩的
    shotScreen(objs[i]['primary'], row, az.index('E'))
    # 生成大人的
    objsSub = objs[i]['secondary']
    for j in range(len(objsSub)):
        alpha = az.index('F')+j
        shotScreen(objsSub[j], row,  alpha)

brguge.close()


input('Press Enter to exit...')




