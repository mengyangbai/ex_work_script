#白
#连接数据库
#依据数据库的时间填补离去时间

import pypyodbc
import io
import sys,os
import time
from dateutil import parser
from datetime import datetime, timedelta

#把数据弄干净一点
def add_two_hour(object):
    if object is not None:
        object = parser.parse(object)
        object += timedelta(hours=2)
        object = object.strftime('%Y-%m-%d %H:%M')
        return object
    return object

def add_depart_time(box_no):
    print("补全"+box_no+"的起飞数据")
    #获取irid和wpos和npos和cdate
    irid = 0
    wpos = 0
    npos = 0
    cdate = "1990-01-01 00:00"
    sqlcommand = """select cr.irid,max(cd.wpos) 
                    from client_rec cr left join check_detail cd on cr.irid = cd.irid 
                    where cr.cnum = '"""
    sqlcommand += box_no
    sqlcommand += "' group by cr.irid;"

    newCursor.execute(sqlcommand)
    
    for row in newCursor:
        irid = row[0]
        wpos = row[1]

    wpos += 1

    sqlcommand = """select cr.irid,max(cd.npos) from client_rec cr left join check_detail cd on cr.irid = cd.irid 
                    where cr.cnum = '"""
    sqlcommand +=box_no+"' and cd.npos <50 group by cr.irid;"
    newCursor.execute(sqlcommand)
    for row in newCursor:
        npos = row[1]   
    sqlcommand ="""select cdate from check_detail where irid = '"""
    sqlcommand+=str(irid)+"""' and npos = '"""
    sqlcommand+=str(npos)+"';"
    newCursor.execute(sqlcommand)
    for row in newCursor:
        cdate = row[0]
    npos += 1
    
    #判断cplace和cinfo
    if "MHT" in box_no:
        cplace = "MELBOURNE - AUSTRALIA"
        cinfo = "Departed Facility in MELBOURNE - 快件已出库"
    elif "SCDA" in box_no or  "SPLA" in box_no or "SPOA" in box_no or "SRYA" in box_no:
        cplace = "SYDNEY - AUSTRALIA"
        cinfo = "Departed Facility in SYDNEY - AUSTRALIA"
    else:
        cplace = "BRISBANE-AUSTRALIA"
        cinfo = "Departed Facility in Brisbane - AUSTRALIA"
    #插入数据
    sqlcommand ="""insert into check_detail(irid,wpos,npos,cplace,cdate,cinfo,cinput,cextra) VALUES('"""
    sqlcommand+= str(irid) + "','"
    sqlcommand+= str(wpos) + "','"
    sqlcommand+= str(npos) + "','"
    sqlcommand+= str(cplace) + "','"
    sqlcommand+= add_two_hour(cdate) + "','"
    sqlcommand+= str(cinfo)+"','张晟','EXCEL');"

    newCursor.execute(sqlcommand)
    newCursor.commit();

    

#解决输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')



