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
import jieba
import io

from tqdm import tqdm
from datetime import datetime, timedelta

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

#sys.stdout = io.TextIOWrapper(sys.stdout.detach(), sys.stdout.encoding, 'replace')

#source_file_dir="resource"
source_file_dir="\\\\192.168.5.214\录单文件\SPV&SHT跑程序"
log_file = "log.txt"

CONFIG_FILE='..\\config\\sql.ini'

#邮件
def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
    ''' Send email
    '''
    
    assert isinstance(send_to, list)
    cf = configparser.ConfigParser()
    cf.read(CONFIG_FILE)
    username = cf.get("EMAIL","user")
    password = cf.get("EMAIL","password")
    
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            msg.attach(part)


    smtp = smtplib.SMTP(server)   
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()


def check(text):
    def __isNum(word):
        if word == '1' or word == '一':
            return False #这里有可能有问题，应该是替换还是怎么样
        pattern = re.compile(r'[0-9]+$')
        if(pattern.match(word)):
            return True
        numList=["一","二","三","四","五","六","七","八","九","十"]
        for tmpWord in word:
            if tmpWord in numList:
                return True
        return False
    
    def __checkNum(Num):
        return __isNum(seg_list[seg_list.index(word)-1])
        
    seg_list = list(jieba.cut(text))#分词
    
    # print("seg_list:"+", ".join(seg_list))
    
    keyword1 = ["组合","*","，"]#这组里面如果存在无脑判断是多品
    for word in keyword1:
        if word in seg_list:
            return True
                
                
    keyword2 = ["盒装","支","盒","袋","包","个","罐","罐装","瓶"] #这组里面如果存在还需判断之前之后的内容
    for word in keyword2:
        if word in seg_list:
            if(__checkNum(word)):
                return True

    keyword3 =["+"]#这组还要check之前会不会是d
    for word in keyword3:
        if word in seg_list:
            if not seg_list[seg_list.index(word)-1].upper() == "D" or not seg_list[seg_list.index(word)-1] == "6" or not seg_list[seg_list.index(word)-1] == "4":
                return True
                
    return False
    

def isBoxNo(boxNo):
    pattern = re.compile(r'[A-Z0-9]+')
    return pattern.match(boxNo)

def connect_to_JERRY():
    cf = configparser.ConfigParser()
    cf.read(CONFIG_FILE)
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
    # ("Start reading "+filename)
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
    
def get_file(files):
    jerry = connect_to_JERRY()
    jerry_cursor = jerry.cursor()
    for single_file in files:
        out_put_file = "output\output_"+single_file
        sqlcommand = """select ib.BOX_NO,im.UNITS_NUMBER,im.PRODUCT_NAME from inventory_basic ib
    left join inventory_merchandise im on ib.id = im.INVENTORY_ID 
    where 

    ib.BOX_NO in """ + readXlsx(single_file,source_file_dir)
        jerry_cursor.execute(sqlcommand)
        
        #如果存在同名文件则删除
        if os.path.isfile(out_put_file):
            os.remove(out_put_file)
        file = xlsxwriter.Workbook(out_put_file)
        
        table = file.add_worksheet('问题箱号')
        table.write(0,0,'问题箱号')
        table.write(0,1,'个数')
        table.write(0,2,'物品描述')
        
        ai_set = set()
        
        n = 1
        for row in jerry_cursor:
            if(check(row[2])):
                ai_set.add((row[0],row[1],row[2]))
                n+=1
                
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
    ib.BOX_NO in """ + readXlsx(single_file,source_file_dir)
        
        jerry_cursor.execute(sqlcommand)
        
        sql_set =set()
        n = 1
        for row in jerry_cursor:
            sql_set.add((row[0],row[1],row[2]))
            n+=1
            
        result_set = ai_set|sql_set
        
        n = 1
        for row in result_set:
            table.write(n,0,row[0])
            table.write(n,1,row[1])
            table.write(n,2,row[2])
            n+=1    
        
        file.close()
    jerry.close()
    with open(log_file, 'a') as f:
        for file in files:
            f.write(file+"\n")
    
    return files
    
def get_all_files():
    tmp_set=set()
    for root, dirs, files in os.walk(source_file_dir):
        for file in files:
            tmp_set.add(file)
    return tmp_set
    
def get_finished_file_list():
    with open(log_file, 'r') as f:
        finished_list = list(set(f.readlines()))
        finished_list = set([m.replace("\n","") for m in finished_list])
    return finished_list
    
if __name__ == '__main__':
    result_set =get_all_files()-get_finished_file_list()
    attach_files=set()
    if len(result_set) != 0:
        attach_files = get_file(result_set)
    
    attach_files=["output\output_"+m for m in attach_files]
    
    #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    send_to = ['lcgroup@ewe.com.au']
    #send_to = ['bain.bai@ewe.com.au']
    subject = 'SPV 多品查询'
    text = 'SPV 多品查询，有问题联系小白： bain.bai@everfast.com.au\n'
    
    if len(attach_files) != 0:
        # pass
        send_mail(send_from,send_to,subject,text,attach_files,server)
