#国内模板
#@author 白
#海豚村AuSydHaituncun AuMelHaituncun
#woolworth 的 So 开头的

import configparser
import pypyodbc
import pymysql
import io
import sys,os
import xlsxwriter
import time
from tqdm import tqdm
from dateutil import parser
from datetime import datetime, timedelta

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

#邮件
def send_mail(send_from, send_to,copy_to, subject, text, files=None,
              server="127.0.0.1"):
    assert isinstance(send_to, list)
    
    username = cf.get("EMAIL","user")
    password = cf.get("EMAIL","password")
    
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Cc'] = COMMASPACE.join(copy_to)
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
    smtp.sendmail(send_from, send_to+copy_to, msg.as_string())
    smtp.close()
    

#把数据弄干净一点
def operate(object):
    if object is not None:
        object = parser.parse(object)
        object = object.strftime('%Y-%m-%d')
        return object
    return object


#解决输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

def exportFinishedExcel(clientname,start_time,end_time):
    filename = "Allen妥投\\"+clientname +" "+start_time+"~"+end_time + ".xlsx"
    #输出
    print("正在输出"+clientname+"的数据到--->"+filename)

    #如果不存在妥投目录则创建
    if not os.path.isdir("Allen妥投"):
        os.mkdir("Allen妥投")
    
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)
    
    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table = file.add_worksheet('报告')
    
    #设置宽度
    table.set_column(0,1,20)
    table.set_column(2,14,10)
    table.write(0,0,'单号')
    table.write(0,1,'MAWB')
    table.write(0,2,'下单日期')
    table.write(0,3,'入库日期')
    table.write(0,4,'出库日期')
    table.write(0,5,'航班起飞日期')
    table.write(0,6,'航班抵达日期')
    table.write(0,7,'清关放行日期')
    table.write(0,8,'送达日期')
    table.write(0,9,'下单-入库时效')
    table.write(0,10,'入库-出库时效')
    table.write(0,11,'出库-到港时效')
    table.write(0,12,'到港-清关时效')
    table.write(0,13,'国内派送时效')
    table.write(0,14,'全程时效')
    
    #先连接sqlserver
    driver = cf.get("EMMIS","driver")
    server = cf.get("EMMIS","server")
    user = cf.get("EMMIS","user")
    password = cf.get("EMMIS","password")
    database = cf.get("EMMIS","database")

    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"

    cnxn = pypyodbc.connect(connection_string)

    cursor = cnxn.cursor()

    #先判断是海豚村还是Woolworths
    if clientname == "AuSydHaituncun":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where (cr.cnum like 'SCDA%' or 
                                cr.cnum like 'SPLA%' or 
                                cr.cnum like 'SPOA%' or 
                                cr.cnum like 'SRYA%') and cr.nstate = 3 and cd.cdate between '"""
    elif clientname == "AuMelHaituncun":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum like 'MHT%' and cr.nstate = 3 and cd.cdate between '"""
    elif clientname == "Woolworths":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum like 'SO%' and cr.nstate = 3 and cd.cdate between '"""        

    sqlcommand += start_time
    sqlcommand += "' and '"
    sqlcommand += end_time
    sqlcommand += "');"
    cursor.execute(sqlcommand)
    
    #进度条
    sys.stdout.write('*' * 10 +"20%"+ '\r')
    sys.stdout.flush()


    sqlcommand = """select Num,
                     MIN(CASE WHEN wpos = 1 THEN cdate END) [Down],
                     MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                     MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                     MIN(CASE WHEN cinfo like 'Departed Facility in%' THEN cdate END) [DEPARTATAUS],
                     MIN(CASE WHEN cinfo = '航班已到达清关口岸' THEN cdate END) [Arrive], 
                     MIN(CASE WHEN cinfo = '清关已完成' THEN cdate END) [ClearanceAccomplished],
                     MAX(CASE WHEN nstate = 3 THEN cdate END) [MissionComplete]
                    FROM #temp1
                    GROUP BY Num
                    ORDER BY Num;"""
    cursor.execute(sqlcommand)
    
    
    #进度条
    sys.stdout.write('*' * 20 +"40%"+ '\r')
    sys.stdout.flush()

    #连接mysql
    server = cf.get("UIS","server")
    port = cf.getint("UIS","port")
    user = cf.get("UIS","user")
    password = cf.get("UIS","password")
    database = cf.get("UIS","database")
    charset = cf.get("UIS","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()

    sqlcommand = """select eo.ordernum,eb.Awb from edb_order eo 
                        LEFT JOIN edb_batch eb on eo.batchId = eb.id 
                        where eo.ordernum in ("""
    sql = []
    
    #录入数据
    n = 1
    for row in cursor:
        table.write(n,0,row[0])
        table.write(n,2,operate(row[1]))
        table.write(n,3,operate(row[2]))
        table.write(n,4,operate(row[3]))
        table.write(n,5,operate(row[4]))
        table.write(n,6,operate(row[5]))
        table.write(n,7,operate(row[6]))
        table.write(n,8,operate(row[7]))
        #下单-入库时效
        if row[1] is not None and row[2] is not None:
            timeofDC = "=D"+str(n+1)+"-C"+str(n+1)
            table.write(n,9,timeofDC)
        #入库-出库时效
        if row[2] is not None and row[3] is not None:
            timeofED = "=E"+str(n+1)+"-D"+str(n+1)
            table.write(n,10,timeofED)
        #出库-到港时效
        if row[3] is not None and row[5] is not None:
            timeofGE = "=G"+str(n+1)+"-E"+str(n+1)
            table.write(n,11,timeofGE)
        #到港-清关时效
        if row[5] is not None and row[6] is not None:
            timeofHG = "=H"+str(n+1)+"-G"+str(n+1)
            table.write(n,12,timeofHG)
        #国内派送时效
        if row[6] is not None and row[7] is not None:
            timeofIH = "=I"+str(n+1)+"-H"+str(n+1)
            table.write(n,13,timeofIH)
        #全程时效
        if row[2] is not None and row[7] is not None:
            timeofID = "=I"+str(n+1)+"-D"+str(n+1)
            table.write(n,14,timeofID)
        sqlcommand += "'"
        sqlcommand += row[0]
        sqlcommand += "',"
        n += 1
        if n / 1000 == 0:
            sqlcommand = sqlcommand.rstrip(',') + ")"
            sqlcommand += " order by eo.ordernum;"
            sql.append(sqlcommand)
            sqlcommand = """select eo.ordernum,eb.Awb from edb_order eo 
                        LEFT JOIN edb_batch eb on eo.batchId = eb.id 
                        where eo.ordernum in ("""
  
        
    #进度条
    sys.stdout.write('*' * 30 +"60%"+ '\r')
    sys.stdout.flush()

    #去尾
    sqlcommand = sqlcommand.rstrip(',') + ")"
    sqlcommand += " order by eo.ordernum;"
    sql.append(sqlcommand)
    
    #进度条
    sys.stdout.write('*' * 40 +"80%"+ '\r')
    sys.stdout.flush()
    
    n = 1
    for sqlcommand in sql:
        cur.execute(sqlcommand)
        for row in cur:
            table.write(n,1,row[1])
            n += 1
    
    

    
    #进度条
    sys.stdout.write('*' * 50 +"100%"+ '\r')
    sys.stdout.flush()
    
    #关闭mysql
    cur.close()
    conn.close()
    
    #关闭sql server
    cursor.close()
    cnxn.close()
    
    #关闭file
    file.close()
    return filename

def exportUnFinishedExcel(clientname):
    filename = "Allen在途\\"+clientname+".xlsx"
    #输出
    print("正在输出"+clientname+"的数据到--->"+filename)

    #如果不存在妥投目录则创建
    if not os.path.isdir("Allen在途"):
        os.mkdir("Allen在途")
    
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)
    
    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table = file.add_worksheet('报告')
    
    #设置宽度
    table.set_column(0,1,20)
    table.set_column(2,14,10)
    table.write(0,0,'单号')
    table.write(0,1,'MAWB')
    table.write(0,2,'下单日期')
    table.write(0,3,'入库日期')
    table.write(0,4,'出库日期')
    table.write(0,5,'航班起飞日期')
    table.write(0,6,'航班抵达日期')
    table.write(0,7,'清关放行日期')
    table.write(0,8,'送达日期')
    table.write(0,9,'下单-入库时效')
    table.write(0,10,'入库-出库时效')
    table.write(0,11,'出库-到港时效')
    table.write(0,12,'到港-清关时效')
    table.write(0,13,'国内派送时效')
    table.write(0,14,'全程时效')
    
    #先连接sqlserver
    driver = cf.get("EMMIS","driver")
    server = cf.get("EMMIS","server")
    user = cf.get("EMMIS","user")
    password = cf.get("EMMIS","password")
    database = cf.get("EMMIS","database")

    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"

    cnxn = pypyodbc.connect(connection_string)

    cursor = cnxn.cursor()

    #先判断是海豚村还是Woolworths
    if clientname == "AuSydHaituncun":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where (cr.cnum like 'SCDA%' or 
                                cr.cnum like 'SPLA%' or 
                                cr.cnum like 'SPOA%' or 
                                cr.cnum like 'SRYA%') and cr.nstate < 3);"""
    elif clientname == "AuMelHaituncun":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum like 'MHT%' and cr.nstate < 3);"""
    elif clientname == "Woolworths":
        sqlcommand = """select cr.cnum Num,cr.nstate,cd.cdate, cd.cinfo,cd.npos,cd.wpos into #temp1 from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum in 
                        (select DISTINCT cr.cnum from client_rec cr LEFT JOIN check_detail cd on cr.irid = cd.irid
                        where cr.cnum like 'SO%' and cr.nstate < 3);"""        

    cursor.execute(sqlcommand)
    
    #进度条
    sys.stdout.write('*' * 10 +"20%"+ '\r')
    sys.stdout.flush()


    sqlcommand = """select Num,
                     MIN(CASE WHEN wpos = 1 THEN cdate END) [Down],
                     MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                     MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                     MIN(CASE WHEN cinfo like 'Departed Facility in%' THEN cdate END) [DEPARTATAUS],
                     MIN(CASE WHEN cinfo = '航班已到达清关口岸' THEN cdate END) [Arrive], 
                     MIN(CASE WHEN cinfo = '清关已完成' THEN cdate END) [ClearanceAccomplished],
                     MAX(CASE WHEN nstate = 3 THEN cdate END) [MissionComplete]
                    FROM #temp1
                    GROUP BY Num
                    ORDER BY Num;"""
    cursor.execute(sqlcommand)
    
    
    #进度条
    sys.stdout.write('*' * 20 +"40%"+ '\r')
    sys.stdout.flush()

    #连接mysql
    server = cf.get("UIS","server")
    port = cf.getint("UIS","port")
    user = cf.get("UIS","user")
    password = cf.get("UIS","password")
    database = cf.get("UIS","database")
    charset = cf.get("UIS","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()

    sqlcommand = """select eo.ordernum,eb.Awb from edb_order eo 
                        LEFT JOIN edb_batch eb on eo.batchId = eb.id 
                        where eo.ordernum in ("""
    
    sql = []
    
    #录入数据
    n = 1
    for row in cursor:
        table.write(n,0,row[0])
        table.write(n,2,operate(row[1]))
        table.write(n,3,operate(row[2]))
        table.write(n,4,operate(row[3]))
        table.write(n,5,operate(row[4]))
        table.write(n,6,operate(row[5]))
        table.write(n,7,operate(row[6]))
        table.write(n,8,operate(row[7]))
        #下单-入库时效
        if row[1] is not None and row[2] is not None:
            timeofDC = "=D"+str(n+1)+"-C"+str(n+1)
            table.write(n,9,timeofDC)
        #入库-出库时效
        if row[2] is not None and row[3] is not None:
            timeofED = "=E"+str(n+1)+"-D"+str(n+1)
            table.write(n,10,timeofED)
        #出库-到港时效
        if row[3] is not None and row[5] is not None:
            timeofGE = "=G"+str(n+1)+"-E"+str(n+1)
            table.write(n,11,timeofGE)
        #到港-清关时效
        if row[5] is not None and row[6] is not None:
            timeofHG = "=H"+str(n+1)+"-G"+str(n+1)
            table.write(n,12,timeofHG)
        #国内派送时效
        if row[6] is not None and row[7] is not None:
            timeofIH = "=I"+str(n+1)+"-H"+str(n+1)
            table.write(n,13,timeofIH)
        #全程时效
        if row[2] is not None and row[7] is not None:
            timeofID = "=I"+str(n+1)+"-D"+str(n+1)
            table.write(n,14,timeofID)
        sqlcommand += "'"
        sqlcommand += row[0]
        sqlcommand += "',"
        n += 1
        if n % 1000 == 0:
            sqlcommand = sqlcommand.rstrip(',') + ")"
            sqlcommand += " order by eo.ordernum;"
            sql.append(sqlcommand)
            sqlcommand = ""
            sqlcommand = """select eo.ordernum,eb.Awb from edb_order eo 
                        LEFT JOIN edb_batch eb on eo.batchId = eb.id 
                        where eo.ordernum in ("""
    
  
        
    #进度条
    sys.stdout.write('*' * 30 +"60%"+ '\r')
    sys.stdout.flush()

    #去尾
    sqlcommand = sqlcommand.rstrip(',') + ")"
    sqlcommand += " order by eo.ordernum;"
    sql.append(sqlcommand)
    
    #进度条
    sys.stdout.write('*' * 40 +"80%"+ '\r')
    sys.stdout.flush()
    
    
    n = 1
    for sqlcommand in sql:
        cur.execute(sqlcommand)
        for row in cur:
            table.write(n,1,row[1])
            n += 1
    
    #进度条
    sys.stdout.write('*' * 50 +"100%"+ '\r')
    sys.stdout.flush()
    
    #关闭mysql
    cur.close()
    conn.close()
    
    #关闭sql server
    cursor.close()
    cnxn.close()
    
    #关闭file
    file.close()
    return filename
    

    
if __name__ == '__main__':
    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')
    files =[]
    filelist = ['AuSydHaituncun','AuMelHaituncun','Woolworths']
    #filelist = ['AuSydHaituncun']
    end_time =datetime.now().strftime('%Y-%m-%d')
    start_time =(datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    for clientname in filelist:
        files.append(exportFinishedExcel(clientname,start_time,end_time))
        files.append(exportUnFinishedExcel(clientname))
    
    
    #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    send_to = ['bain.bai@everfast.com.au']
    copy_to = ['bain.bai@ewe.com.au']
    #copy_to = ['haiyan.xu@ewe.com.au']
    #抄送haiyan.xu@ewe.com.au
    #send_to = ['Allen.li@ewe.com.au']
    #发送Allen.li@ewe.com.au 
    subject = '报表'
    text = '这是自动生成的海豚村与Woolworths的报表从'+start_time+'到'+end_time +'，有问题请联系小白bain.bai@ewe.com.au'
    
    send_mail(send_from,send_to,copy_to,subject,text,files,server)
    

    