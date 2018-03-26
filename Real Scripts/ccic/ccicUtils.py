# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author Bain.Bai
# For implementing basic function of CCIC code scripts

import zipfile
import os,io,functools
import pymysql
import pypyodbc
import configparser
import xlsxwriter
import time
import sys
from datetime import datetime, timedelta

import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

from shutil import copyfile

FIRST_ROW=["快递公司代码","中检包裹溯源码序号","中检包裹溯源码","包裹运单号","包裹收取日期","包裹收取城市","航班号","航班始发城市","航班日期","包裹产品种类","包裹产品件数","包裹重量","快递公司查询链接","空运提单号","单批次包裹的总数量","单批次包裹的总重量","包裹集散仓地址"]
FLIGHT_LOG_FILE="filght_log.txt"
# the script will only send after the local time of AFTER_TIME
AFTER_TIME="2017-10-20 00:00:00.000"
COMPANY_CODE="1001"
#SQL_CONFIG_FILE=r'..\config\sql.ini'
EWE_SEARCH_ADDRESS="https://www.ewe.com.au/track?cno="
SQL_CONFIG_FILE=r'D:\workspace\pythonScripts\python\config\sql.ini'
sys.stdout = io.TextIOWrapper(sys.stdout.detach(), sys.stdout.encoding, 'replace')
dir_dict ={"墨尔本":"\\\\192.168.5.201\MelNormal\\04 物控部\口岸材料","悉尼":"\\\\192.168.5.214\Warehouse\ANDY新发货文件","布里斯班":"\\\\192.168.5.214\Warehouse\布里斯班发货文件"}
output_dir = 'ccic_report'


def get_and_check(airwaybill,city):
    if city in dir_dict:
        #Currently in city
        for root, dirs, files in os.walk(dir_dict[city]):
            for file in files:
                if(airwaybill in file and "Email Copy -" in file):
                    copyfile(os.path.join(root, file),os.path.join(output_dir,file))
                    print("Copying from {} to {}...".format(os.path.join(root, file),os.path.join(output_dir,file)))
                    return True,os.path.join(output_dir,file)
                    
        for anotherCity in dir_dict:
            if city!=anotherCity:
                #Currently in city
                for root, dirs, files in os.walk(dir_dict[anotherCity]):
                    for file in files:
                        if(airwaybill in file and "Email Copy -" in file):
                            copyfile(os.path.join(root, file),os.path.join(output_dir,file))
                            print("{} in wrong city, tell Haiyan, it should be {}".format(airwaybill,anotherCity))
                            return True,os.path.join(output_dir,file)
                        
        print("{} currently not existed".format(airwaybill))
        return False,""
    else:
        for city in dir_dict:
            #Currently in city
            for root, dirs, files in os.walk(dir_dict[city]):
                for file in files:
                    if(airwaybill in file and "Email Copy -" in file):
                        copyfile(os.path.join(root, file),os.path.join(output_dir,file))
                        print("{} in wrong city, tell Haiyan, it should be {}".format(airwaybill,city))
                        return True,os.path.join(output_dir,file)
                       
        print("{} currently not existed".format(airwaybill))
        return False,""
        
#邮件
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

def log(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
        print("Connecting to {}".format(*args))
        return func(*args, **kw)
    return wrapper
    

@log
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


@log
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
    

def export_to_excel(file_list):

    directory = 'ccic_report\\'
    
    def add_finished_flight(flight_list):
        with open(FLIGHT_LOG_FILE, 'a') as f:
            for flight in flight_list:
                f.write(flight+"\n")
    
    flight_list=[]
    output_file_list=[]
    
    for file in file_list:
        filename = directory+file["UisBatchAirwayBill"]+".xlsx"
        # delete if exists
        if os.path.isfile(filename):
            os.remove(filename)
        if len(file["data"]) != 0:
            with xlsxwriter.Workbook(filename) as f:
                table = f.add_worksheet()
                table.write_row('A1',FIRST_ROW)
                n=2
                for oneline in file["data"]:
                    output_line=[]
                    output_line.append(file["company_code"])
                    output_line.append(oneline["ccic_sequence"])
                    output_line.append(oneline["ccic_code"])
                    output_line.append(oneline["boxNo"])
                    output_line.append(oneline["check_weight_time"])
                    output_line.append(file["received_city"])
                    output_line.append(file["flightNo"])
                    output_line.append(file["takeoff_city"])
                    output_line.append(file["DepartedDateTime"])
                    output_line.append(oneline["category"])
                    output_line.append(oneline["number_of_box"])
                    output_line.append(oneline["check_weight"])
                    output_line.append(oneline["search_address"])
                    output_line.append(file["UisBatchAirwayBill"])
                    output_line.append(oneline["total_count"])
                    output_line.append(oneline["total_weight"])
                    output_line.append(file["warehouse_location"])
                    table.write_row('A'+str(n),output_line)
                    n+=1
                
            output_file_list.append(filename)
            print("{} finished.".format(filename))
            result,pdf_file = get_and_check(file["UisBatchAirwayBill"],file["takeoff_city"])
            if result:
                output_file_list.append(pdf_file)
                flight_list.append(file["UisBatchAirwayBill"])
        else:
            flight_list.append(file["UisBatchAirwayBill"])
        
    add_finished_flight(flight_list)
    
    return output_file_list
        
        
            
    
def get_avaiable_Box_info(flight_list):
    
    upc_set=set()
    
    def get_pms_dict(upc_set):
        upc_dict={}
        with connect_to_SQLServer("PMS").cursor() as pms_Cursor:    
            sql_command = '''
            SELECT pro.barcode,ca.NameCn FROM [dbo].[Products] pro left join Category ca 
on pro.CategoryId = ca.Id
where Barcode in ('{}');
            '''.format("','".join(upc_set))
            pms_Cursor.execute(sql_command)
            for row in pms_Cursor:
                if row[1] == "默认":
                    upc_dict[row[0]] = "杂货"
                else:
                    upc_dict[row[0]]= row[1]
        return upc_dict
    
    
    with connect_to_SQLServer("NEWOMS").cursor() as uis_Cursor,connect_to_MYSQL("JERRY").cursor() as jerry_Cursor:
        for flight in flight_list:
            UisBatchAirwayBill = flight["UisBatchAirwayBill"]
            sql_command = '''
            SELECT DISTINCT lc.BoxNo,lc.InboundFirstTime,lc.InboundWeight,
(select count(1) FROM LifeCycle lc left join Batch ba on lc.BatchId = ba.Id
where ba.awb = '{}') as count,
(select sum(lc.InboundWeight) FROM LifeCycle lc left join Batch ba on lc.BatchId = ba.Id
where ba.awb = '{}') as total_weight,
it.RecognizedBarcode
 FROM LifeCycle lc left join Batch ba on lc.BatchId = ba.Id
left join Item it on it.ShipmentId = lc.Id
where ba.awb = '{}';
            '''.format(UisBatchAirwayBill,UisBatchAirwayBill,UisBatchAirwayBill)
            uis_Cursor.execute(sql_command)
            uis_info={}
            uis_box_list=[]
            for row in uis_Cursor:
                tmp={}
                tmp["boxNo"]=row[0]
                # tmp["check_weight_time"]=time.strftime('%Y-%m-%d %H:%M:%S.%f', time.localtime(row[1]))
                tmp["check_weight_time"]=row[1].strftime("%Y-%m-%d")
                tmp["check_weight"]=row[2]
                tmp["total_count"]=row[3]
                tmp["total_weight"]=row[4]
                tmp["search_address"]=EWE_SEARCH_ADDRESS+row[0]
                
                tmp["upc_info"]=set()
                if row[5] is None:
                    pass
                else:
                    tmp["upc_info"].add(row[5])
                    
                upc_set = upc_set|tmp["upc_info"]
                if tmp["boxNo"] not in uis_info:
                    uis_info[tmp["boxNo"]]=tmp
                    uis_box_list.append(tmp["boxNo"])
                else:
                    uis_info[tmp["boxNo"]]["upc_info"]|tmp["upc_info"]
                    
            sql_command = '''
            SELECT CCIC_SEQUENCE,CCIC_CODE,cc.BOX_NO,SUM(im.UNITS_NUMBER) FROM `ccic_code` cc
left join inventory_basic ib on cc.BOX_NO = ib.BOX_NO
left join inventory_merchandise im on im.INVENTORY_ID = ib.ID
 where cc.BOX_NO in ('{}')
 group by cc.BOX_NO;
            '''.format("','".join(uis_box_list))
            jerry_Cursor.execute(sql_command)
            resultData=[]
            for row in jerry_Cursor:
#                if row[2] is not None:
                if row[2].upper() != row[2]:
                    print(row[2])
                    uppertmp=row[2].upper()
                    tmp = uis_info[uppertmp]
                else:
                    tmp = uis_info[row[2]]
                tmp["ccic_sequence"]=row[0]
                tmp["ccic_code"]=row[1]
                tmp["number_of_box"] = row[3]
                resultData.append(tmp)

            flight["data"]=resultData
            
            
    def city_transfer(str):
        cityDict={"MELBOURNE - AUSTRALIA":"墨尔本","SYDNEY - AUSTRALIA":"悉尼","BRISBANE - AUSTRALIA":"布里斯班"}
        if str in cityDict:
            return cityDict[str]
        return str
        
    def flightNo_transfer(str):
        flightNoDict={"默认航班(SYD)":"SYD-CHN","默认航班(MEL)":"MEL-CHN","默认航班(BNE)":"BNE-CHN"}
        if str in flightNoDict:
            return flightNoDict[str]
        return str
        
    def get_warehouse_location(str):
        locationDict={"墨尔本":"14-16 Longford Court, Springvale VIC 3171 Australia","悉尼":"Unit 2，21 Worth Street，Chullora NSW 2190 Australia","布里斯班":"416 Bradman Street, Acacia Ridge QLD 4110 Australia"}
        if str in locationDict:
            return locationDict[str]
        return str
    
    
    upc_dict = get_pms_dict(upc_set)
    with connect_to_SQLServer("TMS").cursor() as tms_Cursor:
        for flight in flight_list:
            sql_command = '''SELECT f.FlightNo,a.City,b.City RealCity FROM [dbo].[Flight] f
left join Airport a on
f.TakeoffAirPortId = a.Id 
left join Airport b on 
f.ActualTakeoffAirPortId = b.Id
 where f.id = {};
            '''.format(flight["FlightId"])
            tms_Cursor.execute(sql_command)
            row = list(tms_Cursor)[0]
            #TMS info
            if row[2] is None:
                flight["received_city"]=city_transfer(row[1])
                flight["flightNo"]=flightNo_transfer(row[0])
                flight["takeoff_city"]=city_transfer(row[1])
            elif row[2] is not None:
                flight["received_city"]=city_transfer(row[1])
                flight["flightNo"]=flightNo_transfer(row[0])
                if row[2] != row[1]:
                    flight["takeoff_city"]=city_transfer(row[1])+"转至"+city_transfer(row[2])
                else:
                    flight["takeoff_city"]=city_transfer(row[2])                
                
            flight["warehouse_location"]=get_warehouse_location(city_transfer(row[1]))
            for row in flight["data"]:
                category_set=set()
                for upc in row["upc_info"]:
                    category_set.add(upc_dict[upc])
                row["category"] = ",".join(category_set)
                
            #other tiny info
            flight["company_code"]=COMPANY_CODE
            
          
    return flight_list
          
                
                
            
    
    
def get_avaiable_flight():

    def get_finished_flight():
        with open(FLIGHT_LOG_FILE, 'r') as f:
            finished_list = list(set(f.readlines()))
            finished_list = [m.replace("\n","") for m in finished_list]
            not_in_string ="'{}'".format("','".join(finished_list))
            return not_in_string
            
    with connect_to_SQLServer("NEWOMS").cursor() as oms_cursor:
        sql_command = '''
        SELECT Awb,FlightId,DepartedDateTime FROM [dbo].[Batch] 
where len(awb) > 0
and awb not in ({}) and DepartedDateTime is not null and FlightId not in ('5','6','7')
and DepartedDateTime >= '{}';
            '''.format(get_finished_flight(),AFTER_TIME)
        oms_cursor.execute(sql_command)
        resultList=[]
        for row in oms_cursor:
            tmp={}
            tmp["UisBatchAirwayBill"]=row[0]
            tmp["FlightId"]=row[1]
            tmp["DepartedDateTime"]=row[2].strftime("%Y-%m-%d")
            resultList.append(tmp)
            
        return resultList


def get_send_files():
    result_files=[]
    for root, dirs, files in os.walk("ccic_report"):
        for file in files:
            if file.endswith(".xlsx"):
                for anofile in files:
                    if file.replace(".xlsx","").strip() in anofile and anofile.endswith(".PDF"):
                        result_files.append(os.path.join(root, file))
                        result_files.append(os.path.join(root, anofile))

    return result_files

def remove_files(files):
    for filename in files:
        if os.path.isfile(filename):
            os.remove(filename)


if __name__=='__main1__':
    remove_files(get_send_files())
    
if __name__=='__main__':
    
    try:
        flight_list = get_avaiable_flight()
        flight_list = get_avaiable_Box_info(flight_list)
        now_time_zip =datetime.now().strftime('%Y-%m-%d')+'.zip'
        output_file_list = export_to_excel(flight_list)
        send_files=get_send_files()
        #发邮件
        server = 'smtp.gmail.com:587'
        send_from = 'bain.bai@everfast.com.au'
        #send_to = ['bain.bai@ewe.com.au','andy.xu@ewe.com.au','ben.feng@ewe.com.au']
        #send_to = ['bain.bai@ewe.com.au']
        send_to = ['haiyan.xu@ewe.com.au','bain.bai@ewe.com.au']
        subject = 'EWE申请溯源码'
        text = 'ewe申请溯源码，有问题请联系小白： bain.bai@everfast.com.au\n'
        #print(send_files)
        if len(send_files) != 0:
            zf = zipfile.ZipFile(now_time_zip, "w", zipfile.zlib.DEFLATED)
            for file in send_files:
                zf.write(file)
            zf.close()
            #send_mail(send_from,send_to,subject,text,now_time_zip,server)
            #remove_files(send_files)
    except Exception as e:
        print(e)
        #发邮件
        server = 'smtp.gmail.com:587'
        send_from = 'bain.bai@everfast.com.au'
        #send_to = ['bain.bai@ewe.com.au','andy.xu@ewe.com.au','ben.feng@ewe.com.au']
        send_to = ['bain.bai@ewe.com.au']
        subject = 'EWE申请溯源码 报警'
        text = e            
        send_mail(send_from,send_to,subject,text,None,server)
