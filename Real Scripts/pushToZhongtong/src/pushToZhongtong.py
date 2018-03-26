# !/usr/bin/env python3
# @Author 白孟阳
# 推给中通的脚本
# -*- coding: utf-8 -*-
import sys 
import pytz
import configparser
import requests
import pymysql
import simplejson
import hashlib
import base64
import time
import pypyodbc
from datetime import datetime
import io

#sys.stdout = io.TextIOWrapper(sys.stdout.detach(), sys.stdout.encoding, 'replace')
sqlConfigFile='D:\\workspace\\pythonScripts\\python\\config\\sql.ini'
#sqlConfigFile='config\\sql.ini'
    
def end():
    print("Closing OMS....")
    omsCursor.close()
    oms.close()
    print("Closing TMS....")
    tmsCursor.close()
    tms.close()
    print("Closing UIS....")
    uisCursor.close()
    uis.close()
    print("Closing EMMIS....")
    emmisCursor.close()
    emmis.close()
    

def connect_to_UISREAL():
    cf = configparser.ConfigParser()
    cf.read(sqlConfigFile)
    #先连接Jerry
    print("Connecting to UISREAL")
    server = cf.get("UISREAL","server")
    port = cf.getint("UISREAL","port")
    user = cf.get("UISREAL","user")
    password = cf.get("UISREAL","password")
    database = cf.get("UISREAL","database")
    charset = cf.get("UISREAL","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    return conn
    
def connect_to_EMMIS():
    cf = configparser.ConfigParser()
    cf.read(sqlConfigFile)
    #先连接EMMIS
    print("Connecting to EMMIS")
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
    return cnxn

def connect_to_OMS():
    cf = configparser.ConfigParser()
    cf.read(sqlConfigFile)
    #先连接OMS
    print("Connecting to OMS")
    driver = cf.get("OMS","driver")
    server = cf.get("OMS","server")
    user = cf.get("OMS","user")
    password = cf.get("OMS","password")
    database = cf.get("OMS","database")
    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"
    cnxn = pypyodbc.connect(connection_string)
    return cnxn
    
def connect_to_TMS():
    cf = configparser.ConfigParser()
    cf.read(sqlConfigFile)
    #先连接TMS
    print("Connecting to TMS")
    driver = cf.get("TMS","driver")
    server = cf.get("TMS","server")
    user = cf.get("TMS","user")
    password = cf.get("TMS","password")
    database = cf.get("TMS","database")
    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"
    tmsCnxn = pypyodbc.connect(connection_string)
    return tmsCnxn    
    
def getcode(data):
    '''根据拼接后的json生成校验码
       XML需要是unicode
    '''
    data = data + "BEB944DED4890720AF"
    temp = data.encode("utf-8")
    md5 = hashlib.md5()
    md5.update(temp)
    md5str = md5.digest()  # 16位
    b64str = base64.b64encode(md5str)
    return b64str

class Zhongtong:
    def __init__(self,row):
        self.logisticsId = row[0]
        self.check_weight_time = row[1]
        self.beater_time = row[2]
        self.check_weight = row[3]
        self.lastStatus = row[4]
    
    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__, 
            sort_keys=True, indent=5)
    
    def push(self):
        currentStatus = self.lastStatus
        if self.check_weight_time is not None:
            optDate = timestamp_to_strtime(self.check_weight_time)
            if self.lastStatus is None or self.lastStatus < 150:
                print("Sending to Zhongtong,in pick up time..OrderNum="+str(self.logisticsId)+"OptDate="+optDate+"Weight="+str(self.check_weight))
                result = pushOneOrderToZhongtong(ZhongtongDto("150",self.logisticsId,optDate,self.check_weight,None,10,None))
                if(result["success"]):
                    self.lastStatus = 150
                else:
                    self.lastStatus = "0"
                    setFinished(self.logisticsId,self.lastStatus,result["msg"])
                    self.lastStatus = 0
                    
            if(self.lastStatus == 150):
                print("Sending to Zhongtong,in warehouse time..OrderNum="+self.logisticsId+"OptDate="+optDate)
                result = pushOneOrderToZhongtong(ZhongtongDto("160",self.logisticsId,optDate,self.check_weight,None,10,None))
                if(result["success"]):
                    self.lastStatus = 160
                else:
                    setFinished(self.logisticsId,self.lastStatus,result["msg"])
                    self.lastStatus = 0

        
        if (self.beater_time is not None and self.lastStatus == 160):
            optDate = timestamp_to_strtime(self.beater_time)
            print("Sending to Zhongtong,in transit time..OrderNum="+self.logisticsId+"OptDate="+optDate)
            result = pushOneOrderToZhongtong(ZhongtongDto("190",self.logisticsId,optDate,self.check_weight,None,10,None))
            if(result["success"]):
                self.lastStatus = 190
            else:
                setFinished(self.logisticsId,self.lastStatus,result["msg"])
                self.lastStatus = 0
            
        if(self.lastStatus == 190):
            flightData = getFlightData(self.logisticsId)
            if "batchId" in flightData:
                if "departTime" in flightData:
                    optDate = flightData["departTime"]
                    print("Sending to Zhongtong,in depart time..OrderNum="+self.logisticsId+"OptDate="+optDate)
                    result = pushOneOrderToZhongtong(ZhongtongDto("220",self.logisticsId,optDate,self.check_weight,flightData["flightName"],10,flightData["transferCode"]))
                    
                    if(result["success"]):
                        self.lastStatus = 220
                    else:
                        setFinished(self.logisticsId,self.lastStatus,result["msg"])
                        self.lastStatus = 0
                    
        if(self.lastStatus == 220):
            flightData = getFlightData(self.logisticsId)
            if "takeDownTime" in flightData:
                optDate = flightData["takeDownTime"]
                print("Sending to Zhongtong,in take down time..OrderNum="+self.logisticsId+"OptDate="+optDate)
                result = pushOneOrderToZhongtong(ZhongtongDto("230",self.logisticsId,optDate,self.check_weight,flightData["flightName"],8,flightData["transferCode"]))
                
                if(result["success"]):
                    self.lastStatus = 230
                else:
                    setFinished(self.logisticsId,self.lastStatus,result["msg"]) 
                    self.lastStatus = 0
        
        if(self.lastStatus == 230):
            flightData = getClearanceData(self.logisticsId)
            if "clearanceStartDate" in flightData:
                optDate = flightData["clearanceStartDate"]
                print("Sending to Zhongtong,in clearance start time..OrderNum="+self.logisticsId+"OptDate="+optDate)
                result = pushOneOrderToZhongtong(ZhongtongDto("360",self.logisticsId,optDate,self.check_weight,None,8,None))
                
                if(result["success"]):
                    self.lastStatus = 360
                else:
                    setFinished(self.logisticsId,self.lastStatus,result["msg"]) 
                    self.lastStatus = 0
        
        if(self.lastStatus == 360):
            flightData = getClearanceData(self.logisticsId)
            if "clearanceFinishDate" in flightData:
                optDate = flightData["clearanceFinishDate"]
                print("Sending to Zhongtong,in clearance finish time..OrderNum="+self.logisticsId+"OptDate="+optDate)
                result = pushOneOrderToZhongtong(ZhongtongDto("371",self.logisticsId,optDate,self.check_weight,None,8,None))
                
                if(result["success"]):
                    self.lastStatus = 371
                else:
                    setFinished(self.logisticsId,self.lastStatus,result["msg"]) 
                    self.lastStatus = 0            
        
        if self.lastStatus != 0 and currentStatus != self.lastStatus:
            setFinishedStatus(self.logisticsId,self.lastStatus)
        
        uis.commit()

def getFlightData(logisticsId):
    result = {}
    sqlcommand = """SELECT batchId from edb_order where beater_ordernum = '"""
    sqlcommand += logisticsId
    sqlcommand += "';"
    uisCursor.execute(sqlcommand)
    for row in uisCursor:
        batchId = row[0]
    if batchId is None:
        return result
    else:
        result['batchId']=batchId
    
    sqlcommand = """SELECT DepartedDateTime,TakeDownDateTime,FlightId,UisbatchAirwayBill FROM UisBatch where UisBatchId = '"""
    sqlcommand += str(batchId)
    sqlcommand += "';"
    omsCursor.execute(sqlcommand)
    for row in omsCursor:
        departTime = row[0].strftime('%Y-%m-%d %H:%M')
        takeDownTime = row[1].strftime('%Y-%m-%d %H:%M')
        flightId = row[2]
        uisbatchAirwayBill = row[3]
        
    if flightId is None:
        return result
    elif departTime is None:
        return result
    
    result["departTime"] = departTime
    result["takeDownTime"] = takeDownTime
    result["transferCode"]=uisbatchAirwayBill
    
    sqlcommand = """SELECT name FROM Flight where id = '"""
    sqlcommand += str(flightId)
    sqlcommand += "';"
    
    tmsCursor.execute(sqlcommand)
    for row in tmsCursor:
        flightName = row[0]
        
    result["flightName"] = flightName
    return result
    
def getClearanceData(logisticsId):
    result = {}
    sqlcommand = """SELECT ordernum FROM edb_order where beater_ordernum = "{}";""".format(logisticsId)
    uisCursor.execute(sqlcommand)
    for row in uisCursor:
        result["boxNo"]=row[0]
        
    sqlcommand = """select cd.cdate from check_detail cd left join client_rec cr on cd.irid= cr.irid
 where cr.cnum = '{}' and cd.cinfo like '%包裹清关中%';""".format(result["boxNo"])
    emmisCursor.execute(sqlcommand)
    for row in emmisCursor:
        result["clearanceStartDate"]=row[0]
    
    sqlcommand = """select top(1) cd.cdate from client_rec cr left join check_detail cd on cd.irid = cr.irid
 where cr.cnum = '{}'
 and  cd.npos >= 100
and not cd.cinfo like '%EWE 进行处理%'
ORDER BY cd.npos;""".format(result["boxNo"])
    emmisCursor.execute(sqlcommand)
    for row in emmisCursor:
        result["clearanceFinishDate"]=row[0]
 
    return result
        
def setFinishedStatus(logisticsId,lastStatus):
    print("Setting status.."+str(logisticsId)+"\tStatus="+str(lastStatus))
    sqlcommand = """SELECT * FROM `edb_transfer_sync` where ordernum = '"""
    sqlcommand += logisticsId
    sqlcommand += "';"
    uisCursor.execute(sqlcommand)
    tmpList = list(uisCursor)
    if len(tmpList) == 0:
        if str(lastStatus) == "371":
            sqlcommand = """INSERT INTO edb_transfer_sync (`ordernum`, `last_status`, `isFinished`, `create_date`, `last_modifed_date`, `remark`) VALUES ('"""
            sqlcommand += logisticsId
            sqlcommand += """', '371', 1, NOW(), NOW(), NULL);"""
        else:
            sqlcommand = """INSERT INTO edb_transfer_sync (`ordernum`, `last_status`, `isFinished`, `create_date`, `last_modifed_date`, `remark`) VALUES ('"""
            sqlcommand += logisticsId
            sqlcommand += "', '"
            sqlcommand += str(lastStatus)
            sqlcommand += """', 0, NOW(), NOW(), NULL);"""
        uisCursor.execute(sqlcommand)
    else:
        if str(lastStatus) == "371":
            sqlcommand = """update edb_transfer_sync set last_status = 371,remark="",isFinished = 1 where ordernum = '"""
            sqlcommand+=logisticsId
            sqlcommand+="""';"""
        else:
            sqlcommand = """update edb_transfer_sync set remark="",last_status = """
            sqlcommand+=str(lastStatus)
            sqlcommand+=""" where ordernum = '"""
            sqlcommand+=logisticsId
            sqlcommand+="""';"""
        uisCursor.execute(sqlcommand)
        
        
 
def setFinished(logisticsId,lastStatus,msg):
    print("Setting status.."+str(logisticsId)+"\tMsg="+str(msg))
    sqlcommand = """SELECT * FROM `edb_transfer_sync` where ordernum = '"""
    sqlcommand += logisticsId
    sqlcommand += "';"
    uisCursor.execute(sqlcommand)
    tmpList = list(uisCursor)
    if len(tmpList) == 0:
        sqlcommand = """INSERT INTO edb_transfer_sync (`ordernum`, `last_status`, `isFinished`, `create_date`, `last_modifed_date`, `remark`) VALUES ('"""
        sqlcommand += logisticsId
        sqlcommand += '''', "'''
        sqlcommand += str(lastStatus)
        sqlcommand += '''", 0, NOW(), NOW(), "'''
        sqlcommand += str(msg)
        sqlcommand += '''");'''
        uisCursor.execute(sqlcommand)
    else:
        sqlcommand = """update edb_transfer_sync set last_status = """
        sqlcommand += str(lastStatus)
        sqlcommand += ''', remark = "'''
        sqlcommand += str(msg)
        sqlcommand += '''" where ordernum = "'''
        sqlcommand += logisticsId
        sqlcommand += '''";'''
        uisCursor.execute(sqlcommand)

class ZhongtongDto:
    #initialise
    def __init__(self,id,logisticsId,optDate,weight,flightCode,zone,transferCode):
        self.id=id
        self.logisticsId=logisticsId
        self.optDate=optDate
        self.optMan="EWE"
        if flightCode is None and weight is not None:
            self.weight=weight
        if flightCode is not None:
            self.flightCode = flightCode
        if transferCode is not None:
            self.transferCode = transferCode
        self.zone=zone
        self.platformSource=1101
        self.warehouseCode="au003"

           
    
    def toJSON(self):
        return json.dumps(self, default=lambda o: o.__dict__, 
            sort_keys=True, indent=5)

def pushOneOrderToZhongtong(zhongtongDto):
    result={}
    data=simplejson.dumps(zhongtongDto, default=lambda obj: obj.__dict__,use_decimal=True)    
    print(data)
    payload = {'data': data, 'msg_type': 'zto.logistics.tracksInfo',"data_digest":getcode(data),"company_id":"AZEWE1012297E107"}
    r = requests.post("https://gjapi.zt-express.com/api/import/init" ,data=payload)
    response = simplejson.loads(r.text)
    if response["status"]:
        result["success"] = True
    elif (response["msg"].encode("utf8") == "不允许此操作(重复称重、已出库或则出库异常)".encode("utf8")):
        result["success"] = True
    elif (response["msg"].encode("utf8") == "不允许此操作(已出库或异常)！".encode("utf8")):
        result["success"] = True
    else:
        result["success"] = False
        result["msg"] = response
        
    return result
    
def getAllZhongTong():
    print("Getting all available data")
    sqlcommand = """SELECT eo.beater_ordernum,eo.check_weight_time,eo.beater_time,
	eo.check_weight,es.last_status FROM `edb_order` eo inner join edb_batch eb on eo.batchid = eb.id
left join edb_transfer_sync es
 on eo.beater_ordernum = es.ordernum
where eo.check_weight_time > 1500182630
and eo.inlandRoute = '天津中通'
and not es.isFinished <=> '1';
"""
    uisCursor.execute(sqlcommand)
    return list(uisCursor)


    
def timestamp_to_strtime(timestamp):
    """将 10 位整数的毫秒时间戳转化成本地普通时间 (字符串格式)
    :param timestamp: 10 位整数的毫秒时间戳 (1456402864242)
    :return: 返回字符串格式 {str}'2016-02-25 20:21:04.242000'
    """
    local_str_time = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')
    return local_str_time

def timestamp_to_datetime(timestamp):
    """将 10 位整数的毫秒时间戳转化成本地普通时间 (datetime 格式)
    :param timestamp: 10 位整数的毫秒时间戳 (1456402864242)
    :return: 返回 datetime 格式 {datetime}2016-02-25 20:21:04.242000
    """
    local_dt_time = datetime.fromtimestamp(timestamp,pytz.timezone('Australia/Sydney'))
    return local_dt_time

if __name__=="__main4__":
    tms=connect_to_TMS()
    tmsCursor=tms.cursor()
    sqlcommand ="""SELECT count(1) FROM flight;"""
    omsCursor.execute(sqlcommand)
    for row in omsCursor:
        print(row)       
    
if __name__=="__main__":
    print("Start")
    uis=connect_to_UISREAL()
    uisCursor=uis.cursor()
    oms=connect_to_OMS()
    omsCursor=oms.cursor()
    tms=connect_to_TMS()
    tmsCursor=tms.cursor()
    emmis=connect_to_EMMIS()
    emmisCursor=emmis.cursor()
    
    # flightData = getClearanceData("120126551316")
    # for key, value in flightData.items() :
        # print (key, value)
                
    

    # optDate = flightData["clearanceStartDate"]
    # result = pushOneOrderToZhongtong(ZhongtongDto("360","120126551279",optDate,None,None,8,None))
    # print(result)
    data = getAllZhongTong()
    for row in data:
        oneOrder = Zhongtong(row)
        oneOrder.push()
    end()
    
if __name__=="__main2__":
    uis=connect_to_UIS()
    uisCursor=uis.cursor()
    sqlcommand="""select count(1) from edb_order;"""
    uisCursor.execute(sqlcommand)
    for row in uisCursor:
        print(row)
    
if __name__=="__main4__":
    uis=connect_to_UIS()
    uisCursor=uis.cursor()
    
    
if __name__=="__main1__":
    test=ZhongtongDto(150,"120002327209","2017-06-21 12:12",1.5,"0")
    data=simplejson.dumps(test, default=lambda obj: obj.__dict__)
    print(data)
    payload = {'data': data, 'msg_type': 'zto.logistics.tracksInfo',"data_digest":getcode(data),"company_id":"AZEWE0952147E106"}
    r = requests.post("http://intltest.zto.cn/api/import/init" ,data=payload)
    print(r.text)