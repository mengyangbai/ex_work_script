# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author Bain.Bai
# For implementing the andy script

import os,io,functools
import pymysql
import configparser
import xlsxwriter
import xlrd

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

SQL_CONFIG_FILE=r'..\config\sql.ini'
SCANNED_FILE='file.txt'
DIR=r".\\"

def connect_to_MYSQL(str):
    ''' connect to MYSQL database
    
    Args:
        str: JERRY,UISREAL,LOCAL_JERRY
    Return:
        Connection, can be used as conn.cursor() and conn.cursor("select count(1) from ...")
    '''
    cf = configparser.ConfigParser()
    cf.read(SQL_CONFIG_FILE)
    server = cf.get(str,"server")
    port = cf.getint(str,"port")
    user = cf.get(str,"user")
    password = cf.get(str,"password")
    database = cf.get(str,"database")
    charset = cf.get(str,"charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    return conn
    
    
def outputfile(filename,data):
    #输出
    print("Exporting "+filename)
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)
        
    with xlsxwriter.Workbook(filename) as file:
        table = file.add_worksheet('report')
        #设置宽度
        table.set_column(0,18,20)
        n = 2
        for row in data:
            table.write_row('A'+str(n),row)
            n=n+1
    print(filename + " exported")
    return filename


def remove_files(files):
    for filename in files:
        if os.path.isfile(filename):
            os.remove(filename)

def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
    ''' Send email
    '''
    
    assert isinstance(send_to, list)
    cf = configparser.ConfigParser()
    cf.read(SQL_CONFIG_FILE)
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
    
def get_available_file():
    file_set=set()
    for root, dirs, files in os.walk(DIR):
        for file in files:
            if file.endswith(".xlsx"):
                file_set.add(file)
    return file_set

def get_finished_file():
    with open(SCANNED_FILE, 'r') as f:
        finished_list = list(set(f.readlines()))
        finished_list = [m.replace("\n","") for m in finished_list]
        #not_in_string ="'{}'".format("','".join(finished_list))
        return set(finished_list)

def add_finished_file(file_list):
    with open(SCANNED_FILE, 'a') as f:
        for filename in file_list:
            f.write(filename+"\n")
            
def make_andy_file(file):
    with xlrd.open_workbook(file) as f:
        original_box_list=[]
        original_data=[]
        sh = f.sheet_by_index(0)
        for row_number in range(sh.nrows):
            #if row_number != 0:
            box_no = sh.cell_value(row_number,0)
            #data1 = sh.cell_value(row_number,1)
            original_box_list.append(box_no)
            tempdata = []
            tempdata.append(box_no)
            #tempdata.append(data1)
            original_data.append(tempdata)
                
        in_string = "'{}'".format("','".join(original_box_list))
        
        box_dict={}
        with connect_to_MYSQL("JERRY").cursor() as jerry_Cursor:
            sql_command = '''select ib.BOX_NO,GROUP_CONCAT(IFNULL(im.BRAND,""),im.PRODUCT_NAME,'*',im.UNITS_NUMBER) as '物品信息'
from inventory_basic ib LEFT JOIN inventory_merchandise im on ib.ID 

 = im.INVENTORY_ID
where ib.BOX_NO in ({}) GROUP BY ib.ID ;'''.format(in_string)

            jerry_Cursor.execute(sql_command)
            for row in jerry_Cursor:
                box_dict[row[0]]=row[1]
                
        for row in original_data:
            if row[0] in box_dict:
                row.append(box_dict[row[0]])
        
        outputfile(file,original_data)        
    
                
        
       
            
if __name__=='__main__':
    
    # try:
        available_files = get_available_file()-get_finished_file()
        for file in available_files:
            make_andy_file(file)
        
        #发邮件
        server = 'smtp.gmail.com:587'
        send_from = 'bain.bai@everfast.com.au'
        #send_to = ['bain.bai@ewe.com.au','andy.xu@ewe.com.au','ben.feng@ewe.com.au']
        send_to = ['bain.bai@ewe.com.au']
        subject = '拉原品名'
        text = 'ewe拉原品名，有问题请联系小白： bain.bai@everfast.com.au\n'
        #print(send_files)
        if len(available_files) != 0:
            send_mail(send_from,send_to,subject,text,available_files,server)
            add_finished_file(available_files)
            #remove_files(available_files)
    # except Exception as e:
        print(e)
        #发邮件
        server = 'smtp.gmail.com:587'
        send_from = 'bain.bai@everfast.com.au'
        #send_to = ['bain.bai@ewe.com.au','andy.xu@ewe.com.au','ben.feng@ewe.com.au']
        send_to = ['bain.bai@ewe.com.au']
        subject = 'EWE andy box 报警'
        # send_mail(send_from,send_to,subject,text,None,server)