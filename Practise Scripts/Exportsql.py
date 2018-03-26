import os
from tqdm import tqdm
import xlrd
import codecs
from datetime import datetime, timedelta

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


#邮件
def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
    assert isinstance(send_to, list)
    
    username = "bain.bai@everfast.com.au"
    password = "Evefast2016"
    
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
    
def writeToSql(filename):
    if os.path.isfile(filename):
        os.remove(filename)   
    output = codecs.open(filename, "w", "utf-8")
    return output

def format_float(num):
    return ('%i' if num == int(num) else '%s') % num    
    
def getSql(barcode,skuname,number):
    basicstr1 = 'INSERT INTO bigorder_type (BARCODE, ITEM_NAME, ITEM_TYPE, ITEM_NUMBER,IS_MIX_BOX,CREATED_TS,CREATED_BY) VALUES ("'
    basicstr2 = '", "'
    basicstr3 = 'w'
    basicstr4 = """1", NOW(), 'bai');"""
    output.write(basicstr1+str(barcode)+basicstr2+skuname+basicstr2+basicstr3+format_float(number)+basicstr2+format_float(number)+basicstr2+basicstr4+"\r\n") 

        
def readXlsx(filename,dir):
    output = filename.replace(" ", "").rstrip(filename[-5:])
    filename = dir + "\\" + filename
    print("开始读取 "+filename)
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)
    for rx in tqdm(range(sh.nrows)):
        var1 = sh.cell_value(rx,0)
        var2 = sh.cell_value(rx,1)
        var3 = sh.cell_value(rx,2)
        if var1.isdigit():
            barcode = var1
            skuname = var2
            number = var3
        else:
            barcode = var2
            skuname = var1
            number = var3
        getSql(barcode,skuname,number)

if __name__ == '__main__':
    
    time = datetime.now().strftime('%Y-%m-%d')
    outputfile = "output"+ time +".sql"
    output = writeToSql(outputfile)
    dir = 'outPutsql'
    try:
        files = os.listdir(dir)
    except FileNotFoundError:
        print("请把待转换的xlsx文件放到程序的outPutsql\目录下！")
        ord(msvcrt.getch())
        quit()
    for file in files:
        if file.endswith(".xlsx"):
            readXlsx(file,dir)
            
    output.close()
    filelist = [outputfile]
    
        #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    #send_to = ['bain.bai@everfast.com.au']
    send_to = ['cen.jia@ewe.com.au']
    subject = 'bigorder 添加'
    text = 'bigorder添加'
    
    send_mail(send_from,send_to,subject,text,filelist,server)
            
        
            