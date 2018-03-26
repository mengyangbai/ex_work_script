#!/usr/bin/env python3
#authoer bain.bai
#墨尔本分店的各种代码
#ASG CMN MAL MBV MCB MCL MCN MDT MGE MGX MHB MJJ MJR MKA MMB MME MMT MMV MNI MOO MQT MSC MSU MTA MTF MTY MWP MWT MYM QVM 
#-f 所有妥投报告，start_time end_time clientname
#-uf 所有在途报告 clientname
#-a 所有报告 以创建时间为标准 start_time end_time clientname

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

SQL_CONFIG_FILE=r'..\config\sql.ini'
FIRST_LINE=['箱号','原单号','货物信息','品名','声称重量','实际重量','收件人城市','发送人城市','置单时间','取件时间','入库时间','出库时长','出库时间','起飞时间','国际运输时长','清关中','清关时长','清关完成','派送时长','妥投/拒收']

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


def connect_to_SQLServer(str):
    ''' connect to MYSQL database
    
    Args:
        str: EMMIS,OMS,TMS
    Return:
        Connection, can be used as conn.cursor() and conn.cursor("select count(1) from ...")
    '''

    cf = configparser.ConfigParser()
    cf.read(SQL_CONFIG_FILE)
    driver = cf.get(str,"driver")
    server = cf.get(str,"server")
    user = cf.get(str,"user")
    password = cf.get(str,"password")
    database = cf.get(str,"database")
    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"
    cnxn = pypyodbc.connect(connection_string)
    return cnxn
    
    
#邮件
def send_mail(send_from, send_to, subject, text, files=None,
              server="127.0.0.1"):
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

#起飞时间调节
#把数据弄干净一点
def add_two_hour(object):
    if object is not None:
        object = parser.parse(object)
        object += timedelta(hours=2)
        object = object.strftime('%Y-%m-%d %H:%M')
        return object
    return object
    
#清关时间添加调节
def add_one_day(object):
    if object is not None:
        object = parser.parse(object)
        object += timedelta(days=1)
        object = object.strftime('%Y-%m-%d %H:%M')
        return object
    return object


#把数据弄干净一点
def operate(object):
    if object is not None:
        object = parser.parse(object)
        object = object.strftime('%Y/%m/%d')
        return object
    return object

#解决输出
#sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

def outputfile(dir,filename,data,start_time,end_time):
    real_file_name = os.path.join(dir,filename)
    #输出
    print("Exporting "+filename)
    #如果不存在妥投目录则创建
    if not os.path.isdir(dir):
        os.mkdir(dir)
        
    #如果存在同名文件则删除
    if os.path.isfile(real_file_name):
        os.remove(real_file_name)
        
    with xlsxwriter.Workbook(real_file_name) as file:
        table = file.add_worksheet('report')
        #设置宽度
        table.set_column(0,18,20)
        report = "关于"+filename+"的在途报告"
        table.write(0,0,report)
        table.write_row('A2',FIRST_LINE)
        n = 3
        for row in data:
            table.write_row('A'+str(n),row)
            n=n+1
    
    print(filename + " exported")
    return real_file_name

def exportIntransitExcel(customer_name,end_time):
    start_time =(datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
    early_time =(datetime.now() - timedelta(days=31)).strftime('%Y-%m-%d')
    result_data=[]
    jerry_dict={}
    jerry_set=set()
    emmis_dict={}
    with connect_to_MYSQL("JERRY").cursor() as jerry_Cursor:
        sql_command = "set session group_concat_max_len = 4096;"
        jerry_Cursor.execute(sql_command)
        sql_command = """select eib.BOX_NO BOX_NO,eob.REFERENCE_NO REFERENCE_NO, GROUP_CONCAT(DISTINCT eim.PRODUCT_NAME SEPARATOR '，') PRODUCT_NAME, GROUP_CONCAT(DISTINCT eim.BRAND SEPARATOR '，') BRAND,
                         eib.WEIGHT DECLARE_WEIGHT,eib.REAL_WEIGHT REAL_WEIGHT ,
                            eoa.CITY RECIEVER_CITY,eoa.SENDER_CITY SENDER_CITY
                            from ewe.inventory_basic eib
                            left join ewe.customer_basic ecb
                            on eib.AGENTPOINT_ID = ecb.ID
                            left join ewe.inventory_merchandise eim
                            on eim.INVENTORY_ID = eib.ID
                            left join ewe.order_basic eob
                            on eib.ORDER_ID = eob.ID
                            left join ewe.order_address eoa
                            on eob.ADDRESS_ID = eoa.ID
                            where ecb.USERNAME = '{}'
and eib.CREATED_DATE>'{}'
and eib.ENABLED_BOX = 'Y' GROUP BY eib.BOX_NO;""".format(customer_name,early_time)
        jerry_Cursor.execute(sql_command)
        for row in jerry_Cursor:
            jerry_set.add(row[0])
            jerry_dict[row[0]]=row
            
    with connect_to_SQLServer("EMMIS").cursor() as emmis_Cursor:
        sql_command ='''select cr.cnum,
MIN(CASE WHEN cd.cinfo = 'Shipment information processed.' THEN cd.cdate END) [SHIPMENT],
                         MIN(CASE WHEN cinfo = 'Package picked up.' or cinfo = 'Picked up by driver.' THEN cdate END) [PICKUP],
                         MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                         MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                         MIN(CASE WHEN cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' or cinfo like 'Departed Facility in%' THEN cdate END) [DEPARTATAUS],
                         MIN(CASE WHEN cinfo = '清关中' or cinfo = '正在清关' or cinfo = '【中国】包裹清关中' THEN cdate END) [Clearance], 
                         MIN(CASE WHEN npos >= 100 THEN cdate END) [ClearanceAccomplished],
                         MAX(CASE WHEN nstate = 3 THEN cdate END) [MissionComplete]
 from client_rec cr left join check_detail cd on cr.irid = cd.irid where cr.cnum in ('{}') and cr.nstate < '3'
group by cr.cnum;
        '''.format("','".join(jerry_set))
        emmis_Cursor.execute(sql_command)
        for row in emmis_Cursor:
            emmis_dict[row[0]]=row
            
    def add_column(row,data):
        if data is not None:
            row.append(data)
        else:
            row.append("")
    
    n=2
    for key in emmis_dict:
        tmprow=[]
        tmprow.extend(jerry_dict[key])
        anotmp_row = emmis_dict[key]    
        add_column(tmprow,operate(anotmp_row[1]))    
        add_column(tmprow,operate(anotmp_row[2]))
        add_column(tmprow,operate(anotmp_row[3]))
        timeofwarehouse = "=IF(OR(ISBLANK(M"+str(n+1)+"),ISBLANK(K"+str(n+1)+")),\"N/A\",M"+str(n+1)+"-K"+str(n+1)+")"
        add_column(tmprow,timeofwarehouse)
        add_column(tmprow,operate(anotmp_row[4]))
        add_column(tmprow,operate(anotmp_row[5]))
        timeofinterflight = "=IF(OR(ISBLANK(P"+str(n+1)+"),ISBLANK(J"+str(n+1)+")),\"N/A\",P"+str(n+1)+"-J"+str(n+1)+")"
        add_column(tmprow,timeofinterflight)
        add_column(tmprow,operate(anotmp_row[6]))
        timeofClearance = "=IF(OR(ISBLANK(R"+str(n+1)+"),ISBLANK(P"+str(n+1)+")),\"N/A\",R"+str(n+1)+"-P"+str(n+1)+")" 
        add_column(tmprow,timeofClearance)
        add_column(tmprow,operate(anotmp_row[7]))
        timeofDeliver = "=IF(OR(ISBLANK(T"+str(n+1)+"),ISBLANK(R"+str(n+1)+")),\"N/A\",T"+str(n+1)+"-R"+str(n+1)+")"
        add_column(tmprow,timeofDeliver) 
        add_column(tmprow,operate(anotmp_row[8]))        
        result_data.append(tmprow)
        n=n+1
        
        
    return outputfile("Intransit",customer_name+".xlsx",result_data,start_time,end_time)
    
    


    
    
if __name__ == '__main__':
    
    file=[]
    end_time =datetime.now().strftime('%Y-%m-%d')
    file.append(exportIntransitExcel("AuMelHaituncun",end_time));
    file.append(exportIntransitExcel("AuSydHaituncun",end_time));
    file.append(exportIntransitExcel("RY",end_time));
    
    #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    send_to = ['bain.bai@everfast.com.au']
    #send_to = ['supplychain@haituncun.com','cen.jia@ewe.com.au','daixiaojing@haituncun.com','hecui@haituncun.com']
    subject = '海豚村自动生成报表'
    text = '这是自动生成的海豚村报表，有问题请联系小白bain.bai@ewe.com.au'
    
    send_mail(send_from,send_to,subject,text,file,server)
