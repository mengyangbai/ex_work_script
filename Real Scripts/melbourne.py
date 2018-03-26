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
def add_two_day(object):
    if object is not None:
        object = parser.parse(object)
        object += timedelta(days=2)
        object = object.strftime('%Y-%m-%d %H:%M')
        return object
    return object

#没有数据时候补全起飞数据
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
    return cdate

#没有数据时候补全清关数据
def add_clearance(box_no):
    print("补全"+box_no+"的清关数据")
    #获取irid和wpos和npos和cdate
    irid = 0
    wpos = 0
    npos = 0
    cdate = "1990-01-01 00:00"
    cplace ="中国"
    cinfo ="正在清关"
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
    sqlcommand +=box_no+"' and cd.npos <100 group by cr.irid;"
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
    

    #插入数据
    sqlcommand ="""insert into check_detail(irid,wpos,npos,cplace,cdate,cinfo,cinput,cextra) VALUES('"""
    sqlcommand+= str(irid) + "','"
    sqlcommand+= str(wpos) + "','"
    sqlcommand+= str(npos) + "','"
    sqlcommand+= str(cplace) + "','"
    sqlcommand+= add_two_day(cdate) + "','"
    sqlcommand+= str(cinfo)+"','张晟','TXT');"

    newCursor.execute(sqlcommand)
    newCursor.commit();
    return cdate

#把数据弄干净一点
def operate(object):
    if object is not None:
        object = parser.parse(object)
        object = object.strftime('%Y/%m/%d')
        return object
    return object

#解决输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

def exportToFinishedExcel(clientname,start_time,end_time):



    filename = "妥投\\"+clientname +" "+start_time+"~"+end_time + ".xlsx"
    
    #输出
    print("正在输出"+clientname+"的数据到--->"+filename)

    #如果不存在妥投目录则创建
    if not os.path.isdir("妥投"):
        os.mkdir("妥投")
    
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)

    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table = file.add_worksheet('报告')
    
    #设置宽度
    table.set_column(0,18,20)
    
    report = "从"+start_time+"到"+end_time+"关于"+clientname+"的妥投报告"
    table.write(0,0,report)
    table.write(1,0,'箱号')
    table.write(1,1,'原单号')
    table.write(1,2,'货物信息')
    table.write(1,3,'品名')
    table.write(1,4,'声称重量')
    table.write(1,5,'实际重量')
    table.write(1,6,'收件人城市')
    table.write(1,7,'发送人城市')
    table.write(1,8,'取件时间')
    table.write(1,9,'入库时间')
    table.write(1,10,'出库时长')
    table.write(1,11,'出库时间')
    table.write(1,12,'起飞时间')
    table.write(1,13,'国际运输时长')
    table.write(1,14,'清关中')
    table.write(1,15,'清关时长')
    table.write(1,16,'清关完成')
    table.write(1,17,'派送时长')
    table.write(1,18,'妥投/拒收')

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

    #先取所有最大pos
    sqlcommand = """select check_detail.irid irid,max(check_detail.npos) npos into #temp1 from check_detail group by check_detail.irid;"""

    cursor.execute(sqlcommand)


    #进度条
    sys.stdout.write('*' * 10 +"20%"+ '\r')
    sys.stdout.flush()



    #先取相应条目
    sqlcommand = """select t.irid irid,t.npos npos into #temp2
                    from #temp1 t LEFT JOIN check_detail 
                    on t.irid = check_detail.irid  
                    and t.npos = check_detail.npos
                    LEFT JOIN client_rec on t.irid = client_rec.irid
                    where left(check_detail.cdate,CHARINDEX(' ', check_detail.cdate)) BETWEEN '"""

    sqlcommand +=start_time
    sqlcommand +="' and '"
    sqlcommand +=end_time
    sqlcommand +="' and (client_rec.nstate >= 3);"
    cursor.execute(sqlcommand)

    sqlcommand = """select client_rec.cnum Num,t.irid irid,check_detail.cdate cdate ,check_detail.cinfo,check_detail.npos npos 
                        into #temp3
                        from #temp2 t left join client_rec on t.irid = client_rec.irid
                        left join check_detail on t.irid = check_detail.irid;"""
    cursor.execute(sqlcommand)

    #rotate
    sqlcommand = """select Num,
                         MIN(CASE WHEN cinfo = 'Package picked up.' or cinfo = 'Picked up by driver.' THEN cdate END) [PICKUP],
                         MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                         MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                         MIN(CASE WHEN cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' or cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' 
                        or cinfo = 'Departed Facility in MELBOURNE - 快件已出库' or cinfo = 'Departed Facility in Brisbane - AUSTRALIA' THEN cdate END) [DEPARTATAUS],
                         MIN(CASE WHEN cinfo = '清关中' or cinfo = '正在清关' or cinfo = '【中国】包裹清关中' THEN cdate END) [Clearance], 
                         MIN(CASE WHEN npos >= 100 THEN cdate END) [ClearanceAccomplished],
                         MAX(CASE WHEN npos >= 100 THEN cdate END) [MissionComplete]
                        into #temp4
                        FROM #temp3
                        GROUP BY Num;"""
                        
    cursor.execute(sqlcommand)


    sqlcommand = "select * from #temp4 order by Num;"

    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 20 +"40%"+ '\r')
    sys.stdout.flush()


    #现在sql数据库里面是要求数据了

    #连接mysql
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()


    sqlcommand = """create TEMPORARY table if not EXISTS temp5(
                        BOX_NO VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""

    cur.execute(sqlcommand)


    #把数据上面的结果数据中的box no插入sql中
    sqlcommand = "insert into temp5 VALUES "
    
    for row in cursor:
        sqlcommand += "('"
        sqlcommand += row[0]
        sqlcommand += "'),"

    #去尾
    sqlcommand = sqlcommand.rstrip(',') + ";"

    cur.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 30 +"60%"+ '\r')
    sys.stdout.flush()

    #在sql获取相应数据
    sqlcommand = "set session group_concat_max_len = 4096;"
    cur.execute(sqlcommand)

    sqlcommand = """select t.BOX_NO BOX_NO,eob.REFERENCE_NO REFERENCE_NO, GROUP_CONCAT(DISTINCT eim.PRODUCT_NAME SEPARATOR '，') PRODUCT_NAME, GROUP_CONCAT(DISTINCT eim.BRAND SEPARATOR '，') BRAND,
                         eib.WEIGHT DECLARE_WEIGHT,eib.REAL_WEIGHT REAL_WEIGHT ,
                            eoa.CITY RECIEVER_CITY,eoa.SENDER_CITY SENDER_CITY
                            from temp5 t
                            inner JOIN ewe.inventory_basic eib
                            on t.BOX_NO = eib.BOX_NO 
                            left join ewe.customer_basic ecb
                            on eib.AGENTPOINT_ID = ecb.ID
                            left join ewe.inventory_merchandise eim
                            on eim.INVENTORY_ID = eib.ID
                            left join ewe.order_basic eob
                            on eib.ORDER_ID = eob.ID
                            left join ewe.order_address eoa
                            on eob.ADDRESS_ID = eoa.ID
                            where ecb.USERNAME = '"""
                            
    sqlcommand += clientname
    sqlcommand += "' GROUP BY t.BOX_NO ORDER BY t.BOX_NO;"

    #录入数据
    cur.execute(sqlcommand)

    #在sql server建立临时表把temp6的boxno反倒进去

    sqlcommand="""create table #temp6(
                        BOX_NO VARCHAR(250) COLLATE Compatibility_199_804_30003
                    );"""

    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 40 +"80%"+ '\r')
    sys.stdout.flush()


    #录入数据并且反向灌入
    n=2
    for row in cur:
        table.write(n,0,row[0])
        table.write(n,1,row[1])
        table.write(n,2,row[2])
        table.write(n,3,row[3])
        table.write(n,4,row[4])
        table.write(n,5,row[5])
        table.write(n,6,row[6])
        table.write(n,7,row[7])
        sqlcommand = "insert into #temp6 VALUES ('"
        sqlcommand += row[0]
        sqlcommand += "');"
        cursor.execute(sqlcommand)
        n+=1

    #进度条
    sys.stdout.write('*' * 45 +"90%"+ '\r')
    sys.stdout.flush()

    #得到的temp4和temp6做inner join

    sqlcommand="""select * into #temp7
                    from #temp6 left join #temp4 on #temp6.BOX_NO = #temp4.Num
                    order by #temp4.Num;"""
                    
    cursor.execute(sqlcommand)    
        
    #录入另一边数据
    #开始录入数据
    #注意由于数据不全，有if判断
    sqlcommand = "select * from #temp7 order by Num;"

    cursor.execute(sqlcommand)
    n=2
    for row2 in cursor:
        if row2[2] is not None:
            table.write(n,8,operate(row2[2]))
        else:
            table.write(n,8,operate(row2[3]))
        table.write(n,9,operate(row2[3]))
        #入库时长
        #=IF(OR(ISBLANK(N6),ISBLANK(P6)),"N\A",P6-N6)
        timeofwarehouse = "=IF(OR(ISBLANK(L"+str(n+1)+"),ISBLANK(J"+str(n+1)+")),\"N/A\",L"+str(n+1)+"-J"+str(n+1)+")"
        #timeofwarehouse = "=K"+str(n+1)+"-I"+str(n+1)
        table.write(n,10,timeofwarehouse)
        table.write(n,11,operate(row2[4]))
        #起飞时间判断和脚本补全
        if row2[5] is not None:
            table.write(n,12,operate(row2[5]))
        elif row2[4] is not None:
            tempDepartTime = add_depart_time(row2[0]);
            table.write(n,12,operate(tempDepartTime))
        else:
            table.write(n,12,operate(row2[5]))
        #国际运输时长
        timeofinterflight = "=IF(OR(ISBLANK(O"+str(n+1)+"),ISBLANK(I"+str(n+1)+")),\"N/A\",O"+str(n+1)+"-I"+str(n+1)+")"
        #timeofinterflight = "=N"+str(n+1)+"-L"+str(n+1)
        table.write(n,13,timeofinterflight)
        #清关时间补全
        if row2[6] is not None:
            table.write(n,14,operate(row2[6]))
        elif row2[5] is not None:
            tempClearanceTime = add_clearance(row2[0])
            table.write(n,14,operate(tempClearanceTime))
        else:
            table.write(n,14,operate(row2[6]))
        #清关时长
        timeofClearance = "=IF(OR(ISBLANK(Q"+str(n+1)+"),ISBLANK(O"+str(n+1)+")),\"N/A\",Q"+str(n+1)+"-O"+str(n+1)+")"    
        #timeofClearance = "=P"+str(n+1)+"-N"+str(n+1)
        table.write(n,15,timeofClearance)
        table.write(n,16,operate(row2[7]))
        #派送时长
        timeofDeliver = "=IF(OR(ISBLANK(S"+str(n+1)+"),ISBLANK(Q"+str(n+1)+")),\"N/A\",S"+str(n+1)+"-Q"+str(n+1)+")"    
        #timeofDeliver = "=R"+str(n+1)+"-P"+str(n+1)
        table.write(n,17,timeofDeliver)
        table.write(n,18,operate(row2[8]))
        n+=1
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
    
def exportToUnfinishedExcel(clientname):
    filename = "在途\\"+clientname + ".xlsx"
        
    print("正在输出"+clientname+"的数据到--->"+filename)

    #如果不存在在途目录则创建
    if not os.path.isdir("在途"):
        os.mkdir("在途")
        
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)
        
    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table = file.add_worksheet('报告')
    
    #设置宽度
    table.set_column(0,18,20)
    
    report = "关于"+clientname+"的在途报告"
    table.write(0,0,report)
    table.write(1,0,'箱号')
    table.write(1,1,'原单号')
    table.write(1,2,'货物信息')
    table.write(1,3,'品名')
    table.write(1,4,'声称重量')
    table.write(1,5,'实际重量')
    table.write(1,6,'收件人城市')
    table.write(1,7,'发送人城市')
    table.write(1,8,'取件时间')
    table.write(1,9,'入库时间')
    table.write(1,10,'出库时长')
    table.write(1,11,'出库时间')
    table.write(1,12,'起飞时间')
    table.write(1,13,'国际运输时长')
    table.write(1,14,'清关中')
    table.write(1,15,'清关时长')
    table.write(1,16,'清关完成')
    table.write(1,17,'派送时长')
    table.write(1,18,'妥投/拒收')

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

    #先取所有最大pos,如果是悉尼的就是除了mht以外的,如果是墨尔本就是mht
    if clientname == "AuSydHaituncun":
        sqlcommand = """select client_rec.cnum from client_rec where client_rec.nstate < 3 and (
                    client_rec.cnum like 'SCDA%' or 
                    client_rec.cnum like 'SPLA%' or 
                    client_rec.cnum like 'SPOA%' or 
                    client_rec.cnum like 'SRYA%');"""
    elif clientname == "AuMelHaituncun":
        sqlcommand = """select client_rec.cnum from client_rec where client_rec.nstate < 3 and (
                    client_rec.cnum like 'MHT%' );"""
    else:
        sqlcommand = """select client_rec.cnum from client_rec where client_rec.nstate < 3;"""

    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 10 +"20%"+ '\r')
    sys.stdout.flush()

    #现在sql数据库里面是要求数据了

    #连接mysql
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()


    sqlcommand = """create TEMPORARY table if not EXISTS temp1(
                        BOX_NO VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""

    cur.execute(sqlcommand)


    #把数据上面的结果数据中的box no插入sql中

    sqlcommand = "insert into temp1 VALUES "
    for row in cursor:    
        sqlcommand += "('"
        sqlcommand += row[0]
        sqlcommand += "'),"

    #去尾
    sqlcommand = sqlcommand.rstrip(',') + ";"
    cur.execute(sqlcommand)
    #进度条
    sys.stdout.write('*' * 20 +"40%"+ '\r')
    sys.stdout.flush()

    #在sql获取相应数据
    sqlcommand = "set session group_concat_max_len = 4096;"
    cur.execute(sqlcommand)

    sqlcommand = """select t.BOX_NO BOX_NO,eob.REFERENCE_NO REFERENCE_NO, GROUP_CONCAT(DISTINCT eim.PRODUCT_NAME SEPARATOR '，') PRODUCT_NAME, GROUP_CONCAT(DISTINCT eim.BRAND SEPARATOR '，') BRAND,
                         eib.WEIGHT DECLARE_WEIGHT,eib.REAL_WEIGHT REAL_WEIGHT ,
                            eoa.CITY RECIEVER_CITY,eoa.SENDER_CITY SENDER_CITY
                            from temp1 t
                            inner JOIN ewe.inventory_basic eib
                            on t.BOX_NO = eib.BOX_NO 
                            left join ewe.customer_basic ecb
                            on eib.AGENTPOINT_ID = ecb.ID
                            left join ewe.inventory_merchandise eim
                            on eim.INVENTORY_ID = eib.ID
                            left join ewe.order_basic eob
                            on eib.ORDER_ID = eob.ID
                            left join ewe.order_address eoa
                            on eob.ADDRESS_ID = eoa.ID
                            where ecb.USERNAME = '"""
                            
    sqlcommand += clientname
    sqlcommand += "'and eib.ENABLED_BOX = 'Y' GROUP BY t.BOX_NO ORDER BY t.BOX_NO;"

    cur.execute(sqlcommand)
    #进度条
    sys.stdout.write('*' * 30 +"60%"+ '\r')
    sys.stdout.flush()

    #在sql server建立临时表把temp6的boxno反倒进去

    sqlcommand="""create table #temp2(
                        BOX_NO VARCHAR(250) COLLATE Compatibility_199_804_30003
                    );"""

    cursor.execute(sqlcommand)


    #录入数据并且反向灌入
    n=2
    for row in cur:
        table.write(n,0,row[0])
        table.write(n,1,row[1])
        table.write(n,2,row[2])
        table.write(n,3,row[3])
        table.write(n,4,row[4])
        table.write(n,5,row[5])
        table.write(n,6,row[6])
        table.write(n,7,row[7])
        sqlcommand = "insert into #temp2 VALUES "
        sqlcommand += "('"
        sqlcommand += row[0]
        sqlcommand += "');"
        cursor.execute(sqlcommand)
        n+=1

    #进度条
    sys.stdout.write('*' * 40 +"80%"+ '\r')
    sys.stdout.flush()

    sqlcommand = """select client_rec.cnum Num,check_detail.cdate cdate ,check_detail.cinfo,check_detail.npos npos,client_rec.nstate nstate
                        into #temp3
                        from #temp2 t left join client_rec on t.BOX_NO = client_rec.cnum
                        left join check_detail on client_rec.irid = check_detail.irid;"""
    cursor.execute(sqlcommand)

    #rotate
    sqlcommand = """select Num,
                         MIN(CASE WHEN cinfo = 'Package picked up.' or cinfo = 'Picked up by driver.' THEN cdate END) [PICKUP],
                         MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                         MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                         MIN(CASE WHEN cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' or cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' 
                        or cinfo = 'Departed Facility in MELBOURNE - 快件已出库' or cinfo = 'Departed Facility in Brisbane - AUSTRALIA' THEN cdate END) [DEPARTATAUS],
                         MIN(CASE WHEN cinfo = '清关中' or cinfo = '正在清关' or cinfo = '【中国】包裹清关中' THEN cdate END) [Clearance], 
                         MIN(CASE WHEN npos >= 100 THEN cdate END) [ClearanceAccomplished],
                         MAX(CASE WHEN nstate = 3 THEN cdate END) [MissionComplete]
                        into #temp4
                        FROM #temp3
                        GROUP BY Num;"""
                        
    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 45 +"90%"+ '\r')
    sys.stdout.flush()
        
    #录入另一边数据
    #开始录入数据
    #注意由于数据不全，有if判断
    sqlcommand = "select * from #temp4 order by Num;"

    cursor.execute(sqlcommand)
    

    n=2
    for row2 in cursor:
        if row2 is not None:
            if row2[2] is not None:
                table.write(n,8,operate(row2[1]))
            else:
                table.write(n,8,operate(row2[2]))
            table.write(n,9,operate(row2[2]))
            #入库时长
            #=IF(OR(ISBLANK(N6),ISBLANK(P6)),"N\A",P6-N6)
            timeofwarehouse = "=IF(OR(ISBLANK(L"+str(n+1)+"),ISBLANK(J"+str(n+1)+")),\"N/A\",L"+str(n+1)+"-J"+str(n+1)+")"
            #timeofwarehouse = "=K"+str(n+1)+"-I"+str(n+1)
            table.write(n,10,timeofwarehouse)
            table.write(n,11,operate(row2[3]))
            table.write(n,12,operate(row2[4]))
            #国际运输时长
            timeofinterflight = "=IF(OR(ISBLANK(Q"+str(n+1)+"),ISBLANK(I"+str(n+1)+")),\"N/A\",Q"+str(n+1)+"-I"+str(n+1)+")"
            #timeofinterflight = "=N"+str(n+1)+"-L"+str(n+1)
            table.write(n,13,timeofinterflight)
            table.write(n,14,operate(row2[5]))
            #清关时长
            timeofClearance = "=IF(OR(ISBLANK(Q"+str(n+1)+"),ISBLANK(O"+str(n+1)+")),\"N/A\",Q"+str(n+1)+"-O"+str(n+1)+")"    
            #timeofClearance = "=P"+str(n+1)+"-N"+str(n+1)
            table.write(n,15,timeofClearance)
            table.write(n,16,operate(row2[6]))
            #派送时长
            timeofDeliver = "=IF(OR(ISBLANK(S"+str(n+1)+"),ISBLANK(Q"+str(n+1)+")),\"N/A\",S"+str(n+1)+"-Q"+str(n+1)+")"    
            #timeofDeliver = "=R"+str(n+1)+"-P"+str(n+1)
            table.write(n,17,timeofDeliver)
            table.write(n,18,operate(row2[7]))
        n+=1



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

def exportToExcel(clientname,start_time,end_time):

    filename = "所有\\"+clientname +" "+start_time+"~"+end_time + ".xlsx"
    
    #输出
    print("正在输出"+clientname+"的数据到--->"+filename)

    #如果不存在在途目录则创建
    if not os.path.isdir("所有"):
        os.mkdir("所有")
    
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)    

    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table = file.add_worksheet('报告')
    
    #设置宽度
    table.set_column(0,18,20)
    
    report = "从(创建时间)"+start_time+"到"+end_time+"关于"+clientname+"的所有报告"
    table.write(0,0,report)
    table.write(1,0,'箱号')
    table.write(1,1,'货物信息')
    table.write(1,2,'品名')
    table.write(1,3,'声称重量')
    table.write(1,4,'实际重量')
    table.write(1,5,'收件人城市')
    table.write(1,6,'发送人城市')
    table.write(1,7,'取件时间')
    table.write(1,8,'入库时间')
    table.write(1,9,'出库时长')
    table.write(1,10,'出库时间')
    table.write(1,11,'起飞时间')
    table.write(1,12,'国际运输时长')
    table.write(1,13,'清关中')
    table.write(1,14,'清关时长')
    table.write(1,15,'清关完成')
    table.write(1,16,'派送时长')
    table.write(1,17,'妥投/拒收')

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




    #先取相应条目
    sqlcommand = """select client_rec.irid irid,client_rec.cnum Num into #temp1
                    from client_rec LEFT JOIN check_detail 
                    on client_rec.irid = check_detail.irid
                    where client_rec.dsysdate BETWEEN '"""

    sqlcommand +=start_time
    sqlcommand +="' and '"
    sqlcommand +=end_time
    sqlcommand +="';"
    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 10 +"20%"+ '\r')
    sys.stdout.flush()

    sqlcommand = """select client_rec.cnum Num,t.irid irid,check_detail.cdate cdate ,check_detail.cinfo,check_detail.npos npos,client_rec.nstate nstate
                        into #temp2
                        from #temp1 t left join client_rec on t.irid = client_rec.irid
                        left join check_detail on t.irid = check_detail.irid;"""
    cursor.execute(sqlcommand)

    #rotate
    sqlcommand = """select Num,
                         MIN(CASE WHEN cinfo = 'Package picked up.' or cinfo = 'Picked up by driver.' THEN cdate END) [PICKUP],
                         MIN(CASE WHEN cinfo = 'Package arrived at warehouse.' THEN cdate END) [ARRIVEATWAREHOUSE],
                         MIN(CASE WHEN cinfo = 'In transit to airport.' THEN cdate END) [TRANSITTOAIRPORT],
                         MIN(CASE WHEN cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' or cinfo = 'Departed Facility in SYDNEY - AUSTRALIA' 
                        or cinfo = 'Departed Facility in MELBOURNE - 快件已出库' or cinfo = 'Departed Facility in Brisbane - AUSTRALIA' THEN cdate END) [DEPARTATAUS],
                         MIN(CASE WHEN cinfo = '清关中' or cinfo = '正在清关' or cinfo = '【中国】包裹清关中' THEN cdate END) [Clearance], 
                         MIN(CASE WHEN npos >= 100 THEN cdate END) [ClearanceAccomplished],
                         MAX(CASE WHEN nstate = 3 THEN cdate END) [MissionComplete]
                        into #temp3
                        FROM #temp2
                        GROUP BY Num;"""
                        
    cursor.execute(sqlcommand)


    sqlcommand = "select * from #temp3 order by Num;"

    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 20 +"40%"+ '\r')
    sys.stdout.flush()


    #现在sql数据库里面是要求数据了

    #现在sql数据库里面是要求数据了

    #连接mysql
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()


    sqlcommand = """create TEMPORARY table if not EXISTS temp4(
                        BOX_NO VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""

    cur.execute(sqlcommand)


    #把数据上面的结果数据中的box no插入sql中
    sqlcommand = "insert into temp4 VALUES "
    for row in cursor:
        sqlcommand += "('"
        sqlcommand += row[0]
        sqlcommand += "'),"

    #去尾
    sqlcommand = sqlcommand.rstrip(',') + ";"

    cur.execute(sqlcommand)


    #进度条
    sys.stdout.write('*' * 30 +"60%"+ '\r')
    sys.stdout.flush()

    #在sql获取相应数据
    sqlcommand = "set session group_concat_max_len = 4096;"
    cur.execute(sqlcommand)

    sqlcommand = """select t.BOX_NO BOX_NO, GROUP_CONCAT(DISTINCT eim.PRODUCT_NAME SEPARATOR '，') PRODUCT_NAME, GROUP_CONCAT(DISTINCT eim.BRAND SEPARATOR '，') BRAND,
                         eib.WEIGHT DECLARE_WEIGHT,eib.REAL_WEIGHT REAL_WEIGHT ,
                            eoa.CITY RECIEVER_CITY,eoa.SENDER_CITY SENDER_CITY
                            from temp4 t
                            inner JOIN ewe.inventory_basic eib
                            on t.BOX_NO = eib.BOX_NO 
                            left join ewe.customer_basic ecb
                            on eib.AGENTPOINT_ID = ecb.ID
                            left join ewe.inventory_merchandise eim
                            on eim.INVENTORY_ID = eib.ID
                            left join ewe.order_basic eob
                            on eib.ORDER_ID = eob.ID
                            left join ewe.order_address eoa
                            on eob.ADDRESS_ID = eoa.ID
                            where ecb.USERNAME = '"""
                            
    sqlcommand += clientname
    sqlcommand += "'and eib.ENABLED_BOX = 'Y' GROUP BY t.BOX_NO ORDER BY t.BOX_NO;"

    #录入数据
    cur.execute(sqlcommand)

    #在sql server建立临时表把temp6的boxno反倒进去

    sqlcommand="""create table #temp5(
                        BOX_NO VARCHAR(250) COLLATE Compatibility_199_804_30003
                    );"""

    cursor.execute(sqlcommand)

    #进度条
    sys.stdout.write('*' * 40 +"80%"+ '\r')
    sys.stdout.flush()


    #录入数据并且反向灌入
    n=2
    for row in cur:
        table.write(n,0,row[0])
        table.write(n,1,row[1])
        table.write(n,2,row[2])
        table.write(n,3,row[3])
        table.write(n,4,row[4])
        table.write(n,5,row[5])
        table.write(n,6,row[6])
        sqlcommand = "insert into #temp5 VALUES ('"
        sqlcommand += row[0]
        sqlcommand += "');"
        cursor.execute(sqlcommand)
        n+=1

    #进度条
    sys.stdout.write('*' * 45 +"90%"+ '\r')
    sys.stdout.flush()

    #得到的temp4和temp6做inner join

    sqlcommand="""select * into #temp6
                    from #temp5 left join #temp3 on #temp5.BOX_NO = #temp3.Num
                    order by #temp3.Num;"""
                    
    cursor.execute(sqlcommand)    
        
    #录入另一边数据
    #开始录入数据
    #注意由于数据不全，有if判断
    sqlcommand = "select * from #temp6 order by Num;"

    cursor.execute(sqlcommand)
    n=2
    for row2 in cursor:
        if row2[2] is not None:
            table.write(n,7,operate(row2[2]))
        else:
            table.write(n,7,operate(row2[3]))
        table.write(n,8,operate(row2[3]))
        #入库时长
        #=IF(OR(ISBLANK(N6),ISBLANK(P6)),"N\A",P6-N6)
        timeofwarehouse = "=IF(OR(ISBLANK(K"+str(n+1)+"),ISBLANK(I"+str(n+1)+")),\"N/A\",K"+str(n+1)+"-I"+str(n+1)+")"
        #timeofwarehouse = "=K"+str(n+1)+"-I"+str(n+1)
        table.write(n,9,timeofwarehouse)
        table.write(n,10,operate(row2[4]))
        table.write(n,11,operate(row2[5]))
        #国际运输时长
        timeofinterflight = "=IF(OR(ISBLANK(N"+str(n+1)+"),ISBLANK(H"+str(n+1)+")),\"N/A\",N"+str(n+1)+"-H"+str(n+1)+")"
        #timeofinterflight = "=N"+str(n+1)+"-L"+str(n+1)
        table.write(n,12,timeofinterflight)
        table.write(n,13,operate(row2[6]))
        #清关时长
        timeofClearance = "=IF(OR(ISBLANK(P"+str(n+1)+"),ISBLANK(N"+str(n+1)+")),\"N/A\",P"+str(n+1)+"-N"+str(n+1)+")"    
        #timeofClearance = "=P"+str(n+1)+"-N"+str(n+1)
        table.write(n,14,timeofClearance)
        table.write(n,15,operate(row2[7]))
        #派送时长
        timeofDeliver = "=IF(OR(ISBLANK(R"+str(n+1)+"),ISBLANK(P"+str(n+1)+")),\"N/A\",R"+str(n+1)+"-P"+str(n+1)+")"    
        #timeofDeliver = "=R"+str(n+1)+"-P"+str(n+1)
        table.write(n,16,timeofDeliver)
        table.write(n,17,operate(row2[8]))
        n+=1
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
    file =[]
    
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
    newCnxn = pypyodbc.connect(connection_string)

    newCursor = newCnxn.cursor()
    
    end_time =datetime.now().strftime('%Y-%m-%d')
    start_time =(datetime.now() - timedelta(days=10)).strftime('%Y-%m-%d')
    #filelist = ["ASG","CMN","DLD","MAL","MBV","MCB","MCL","MCN","MDT","MGE","MGX","MHB","MJJ","MJR","MKA","MMB","MME","MMT","MMV","MNI","MOO","MQT","MSC","MSU","MTA","MTF","MTY","MWP","MWT","MYM","QVM"]
    filelist = ["ASG"]
    for tmp in filelist:
        file.append(exportToExcel(tmp,start_time,end_time))

    
    #发邮件
    server = 'smtp.gmail.com:587'
    send_from = 'bain.bai@everfast.com.au'
    send_to = ['bain.bai@everfast.com.au']
    copy_to = ['bain.bai@ewe.com.au']
    #send_to = ['chris.mel@ewe.com.au','simon.xu@ewe.com.au']
    #copy_to = ['peter.yang@ewe.com.au','cen.jia@ewe.com.au','shanshan.yang@ewe.com.au','info.mel@ewe.com.au']
    subject = '墨尔本自动生成报表'
    text = '这是自动生成的墨尔本报表从'+start_time+'到'+end_time +'，有问题请联系小白bain.bai@ewe.com.au'
    
    send_mail(send_from,send_to,copy_to,subject,text,file,server)

    
    #关闭
    newCursor.close()
    newCnxn.close()
