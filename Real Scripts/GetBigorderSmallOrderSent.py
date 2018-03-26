import os
from tqdm import tqdm
import xlrd
import codecs
from datetime import datetime, timedelta
import pymysql
import configparser
import xlsxwriter
import msvcrt
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
    
    print("开始发送邮件")
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
    

def connect_to_JERRY():

    cf = configparser.ConfigParser()
    cf.read('..\\config\\sql.ini')
    #先连接Jerry
    print("连接到JERRY")
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    return conn


if __name__ == '__main__':
    
    time = datetime.now().strftime('%Y-%m-%d')
    outputfile = "output"+ time +".xlsx"
    file = xlsxwriter.Workbook(outputfile)
    table = file.add_worksheet('报告')
    table.write(0,0,"大订单箱号")
    table.write(0,1,'小订单订单号')
    table.write(0,2,'小订单订单号')
    table.write(0,3,'序号')
    table.write(0,4,'是否为ECI')
    table.write(0,5,'核重时间')
    
    jerry = connect_to_JERRY()
    jerry_cursor = jerry.cursor()
    sqlcommand = """select bo.BIG_BOX_NO,ob.ORDER_NO,ib.BOX_NO,ob.BIGORDER_INDEX,ob.IS_ECI,DATE_FORMAT(DATE_ADD(bo.CREATED_DATE,INTERVAL 10 hour),'%Y-%m-%d'),ib.STORAGE_TS from bigorder bo LEFT JOIN order_basic ob
                    on bo.id = ob.BIGORDER_ID
                    left join inventory_basic ib on ob.ID = ib.ORDER_ID;"""

    jerry_cursor.execute(sqlcommand)
    n=1
    print("开始生成数据")
    for row in tqdm(jerry_cursor):
        table.write(n,0,row[0])
        table.write(n,1,row[1])
        table.write(n,2,row[2])
        table.write(n,3,row[3])
        table.write(n,4,row[4])
        table.write(n,5,row[5])
        n+=1
            
    file.close()
    filelist = [outputfile]
    
        #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    send_to = ['boyka.zhang@ewe.com.au']
    #send_to = ['vincent.wang@ewe.com.au']
    #send_to = ['cen.jia@ewe.com.au']
    subject = 'bigorder与smallorder对应的表'
    text = '此列表为脚本自动生成，有问题联系小白bain.bai@everfast.com.au'
    
    send_mail(send_from,send_to,subject,text,filelist,server)
    print("按任意键退出")     
    msvcrt.getch()
            