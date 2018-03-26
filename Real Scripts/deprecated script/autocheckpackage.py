# !/usr/bin/env python3
# @authoer bain.bai
# 给andy自动分箱的脚本

import re
import configparser
import pymysql
import sys,os
import xlsxwriter
import locale
import xlrd

from tqdm import tqdm
from datetime import datetime, timedelta

def isBoxNo(boxNo):
    pattern = re.compile(r'[A-Z0-9]+')
    return pattern.match(boxNo)

def connect_to_JERRY():
    cf = configparser.ConfigParser()
    cf.read('..\\config\\sql.ini')
    #先连接Jerry
    print("Connecting to JERRY")
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    return conn
    
def readXlsx(filename,dir):
    output = filename.replace(" ", "").rstrip(filename[-5:])
    filename = dir + "\\" + filename
    print("开始读取 "+filename)
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)
    res="("
    for rx in tqdm(range(sh.nrows)):
        boxNo = sh.cell_value(rx,1)
        if isBoxNo(boxNo):
            res = res + "'"
            res = res + boxNo
            res = res + "',"
    res = res[:-1] +");"
    return res
    
if __name__ == '__main__':
    jerry = connect_to_JERRY()
    jerry_cursor = jerry.cursor()
    sqlcommand = """select ib.BOX_NO,im.UNITS_NUMBER,im.PRODUCT_NAME from inventory_basic ib
left join inventory_merchandise im on ib.id = im.INVENTORY_ID 
where 
(
(im.PRODUCT_NAME like "%【%") or 
(im.PRODUCT_NAME like "%+%" and im.PRODUCT_NAME not like "%D+%" and im.PRODUCT_NAME not REGEXP '[0-9]+[[.plus-sign.]]') or
(im.PRODUCT_NAME like "%*%") or
(im.PRODUCT_NAME like "%盒装%") or
(im.PRODUCT_NAME like "%支%") or
(im.PRODUCT_NAME like "%袋%") or
(im.PRODUCT_NAME like "%包%") or 
(im.PRODUCT_NAME like "%，%") or 
(im.PRODUCT_NAME like "%个%") or 
(im.PRODUCT_NAME like "%三%" and im.PRODUCT_NAME not like "%三段%") or
(im.PRODUCT_NAME like "%罐%" and im.PRODUCT_NAME not like "%罐装%")
)
AND
ib.BOX_NO in """ + readXlsx("test.xlsx","test")
    print(sqlcommand)
    jerry_cursor.execute(sqlcommand)
    
    #如果存在同名文件则删除
    if os.path.isfile("output.xlsx"):
        os.remove("output.xlsx")
    file = xlsxwriter.Workbook("output.xlsx")
    
    table = file.add_worksheet('问题箱号')
    table.write(0,0,'问题箱号')
    table.write(0,1,'个数')
    table.write(0,2,'物品描述')
    
    n = 1
    for row in jerry_cursor:
        table.write(n,0,row[0])
        table.write(n,1,row[1])
        table.write(n,2,row[2])
        n+=1
    
    file.close()
    jerry.close()