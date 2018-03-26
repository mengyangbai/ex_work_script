# !/usr/bin/env python3
# @Author 白孟阳
# 把ECWMS的excel转换成runbow的两个excel
import xlrd
import xlwt
import msvcrt
import os
import simplejson
import copy
import operator
from tqdm import tqdm
from datetime import datetime
import configparser
import pymysql
import xlsxwriter

cainiaoList = ["贝拉米海外旗舰店","woolworths海外旗舰店","springleaf海外旗舰店","naturesorganics旗舰店","metcash官方海外旗舰店","jessicassuitcase海外旗舰","evanstate海外旗舰店","chemistwarehouse海外旗舰","babycare海外专营店"]
jerryList = ["ABM","ATA","EA","GOME","JSS","KJLM","RY","SBC","TopWarehouse","DigitalJungle","ONC"]
ECWMSJERRYDict={"ONC":"Concord-Oz Natural Care","DigitalJungle":"sydney-DIGITAL JUNGLE","ABM":"Sydney-ACCESS BRAND MANAGEMENT","ATA":"Smithfield-Austvita","EA":"LandCove-ATW-Aotao","GOME":"ADS-SGM国美","JSS":"JSS","KJLM":"ADS-SDA跨境联盟","RY":"RY","SBC":"ADS-SBC-Babycare","TopWarehouse":"MIO仓储"}

sqlConfigFile='D:\\workspace\\pythonScripts\\python\\config\\sql.ini'

input_dir = 'ECWMS'
output_dir_ASN = 'RUNBOW-ASN'
output_dir_Located = 'RUNBOW-LOCATED'
ecwmsDict = ["LotATT04","UPC","Descr C","Descr E","LOC","ID","InvQty","AllQty","HoldQty","AvblQty","Production Date","Expiration Date","Inbount Date"]
ASN_First_page_first_line=["外部入库单号","预入库单类型","预入库日期","客户","仓库名称","备注"]
ASN_Second_page_first_line=["外部入库单号","SKU","产品名称","UPC","预收数量","单位","规格","批次号","托号","长","宽","高","过期日期"]
UpShelf_FirstLine=["入库单号","客户名称","外部单号","UPC","SKU","收货单行号","货品名称","货品类型","实际数量","仓库","库区","库位","批次号","托号","单位","规格","备注","生产日期","过期日期","创建时间","创建人"]
problem_first_line=["问题","原名","ecwmsupc","jerry名","jerrySku1","jerryUpc1","jerryQty1","jerrySku2","jerryUpc2","jerryQty2","jerrySku3","jerryUpc3","jerryQty3"]
ECWMS_to_RUNBOW_Dict={"ONC":"ONC","DigitalJungle":"DigitalJungle","贝拉米海外旗舰店":"cainiao_bellamy","woolworths海外旗舰店":"cainiao_wws","springleaf海外旗舰店":"cainiao_springleaf","naturesorganics旗舰店":"cainiao_naturesorganics","metcash官方海外旗舰店":"cainiao_metcash","jessicassuitcase海外旗舰":"cainiao_jss","evanstate海外旗舰店":"cainiao_evanstate","chemistwarehouse海外旗舰":"cainiao_cw","babycare海外专营店":"cainiao_babycare","ABM":"SBR","ATA":"SMF","BUBL":"BUBL","EA":"EA","GOME":"SGM","JSS":"JSS","KJLM":"SDA","LAMO":"Lamos","RY":"RP-1","SBC":"SBC","TopWarehouse":"NCJ",}

#存储所有问题数据的字典
problemDict={}
problemList=[]
#存储所有库存数据的字典
allQtyDict={}

users=[]

def checkArea(str):
    if str == "DamageItems":
        str = "FL-Damage-1"
    elif str == "ExpiredItems":
        str = "FL-Expired-1"
    result = "Damage and expire"
    # else:
        # print(str)
        # result = "注意"
    return result,str

def initial():
    print("初始化开始")
    sqlcommand="""CREATE TEMPORARY TABLE IF NOT EXISTS table1  as (SELECT come.ITEM_ID as sku,come.LENGTH as length,come.WIDTH as width,come.HEIGHT as height,come.`NAME` as chinese_name,(SELECT cur.SUBSCRIBER_NICK from cn_user cur where cur.USER_ID = come.OWNER_USER_ID LIMIT 1) as username,come.BAR_CODE as upc,
		(come.NUM-IFNULL(outn.NUM,0)-IFNULL(pandian.QUANTITY,0)) AS 'qty'
 	FROM
		(select comm.ITEM_ID,notify.STORE_CODE,comm.OWNER_USER_ID,comm.`NAME`,comm.LENGTH,comm.WIDTH,comm.HEIGHT,comm.BAR_CODE,inv.INVENTORYTYPE,SUM(inv.QUANTITY) AS 'NUM' from cn_stock_tally_order_item_inv inv
		LEFT JOIN cn_stock_tally_order_item item on inv.FK_ID = item.ID
		LEFT JOIN cn_stock_tally_info tally on item.FK_ID = tally.ID
		LEFT JOIN cn_commodity_info comm ON comm.ITEM_ID = item.ITEMID
		LEFT JOIN cn_stock_in_order_notify notify ON notify.ORDER_CODE = tally.ORDERCODE
	WHERE  tally.`STATUS` = 1 AND tally.UPDATED_TS < now()
		GROUP BY comm.ITEM_ID,notify.STORE_CODE,inv.INVENTORYTYPE) as come
		LEFT JOIN
		(
		select
tempout.ITEM_ID,
tempout.STORE_CODE,
tempout.`NAME`,
tempout.BAR_CODE,
SUM(tempout.ITEM_QUANTITY) as 'NUM',
tempout.INVENTORY_TYPE
from
(
SELECT
		com.ITEM_ID,
		ord.STORE_CODE,
		com.`NAME`,
		com.BAR_CODE,
		citem.ITEM_QUANTITY,
		citem.INVENTORY_TYPE
	FROM
		cn_package_info_item citem
	LEFT JOIN cn_package_info bag ON citem.PACKAGE_INFO_ID = bag.ID
	LEFT JOIN cn_consign_order_notify ord ON bag.CONSIGN_ORDER_NOTIFY_ID = ord.ID
	LEFT JOIN cn_commodity_info com ON citem.ITEM_ID = com.ITEM_ID
	WHERE
		ord. STATUS = 10
	AND ord.LAST_MODIFIED_DATE < now()

UNION ALL
select
		ocom.ITEM_ID,
		oord.STORE_CODE,
		ocom.`NAME`,
		ocom.BAR_CODE,
		opagi.ITEM_QUANTITY,
		opagi.INVENTORY_TYPE
from cn_stock_out_package_info_item opagi
LEFT JOIN cn_stock_out_package_info opag ON opagi.PACKAGE_INFO_ID = opag.ID
LEFT JOIN cn_stock_out_order_notify oord ON opag.STOCK_OUT_ORDER_NOTIFY_ID=oord.ID
LEFT JOIN cn_commodity_info ocom ON opagi.ITEM_ID = ocom.ITEM_ID
where oord.`STATUS` = 10
and oord.ORDER_CODE not in ('LBX012920758554794')
and oord.LAST_MODIFIED_DATE < NOW()
) as tempout
GROUP BY tempout.ITEM_ID,tempout.STORE_CODE,tempout.INVENTORY_TYPE
		) as outn
		ON come.ITEM_ID=outn.ITEM_ID and come.STORE_CODE = outn.STORE_CODE AND come.INVENTORYTYPE = outn.INVENTORY_TYPE
LEFT JOIN

	(
		SELECT pdl.STORE_CODE,pdl.ITEM_ID,pdl.INVENTORY_TYPE,sum(pdl.QUANTITY)AS 'QUANTITY' FROM (
			SELECT ic.STORE_CODE,cd.ITEM_ID,
			(CASE WHEN ic.ORDER_TYPE=701 THEN cd.QUANTITY ELSE cd.QUANTITY-cd.QUANTITY-cd.QUANTITY END) AS 'QUANTITY' ,
			ic.UPDATED_TS,cd.INVENTORY_TYPE,
			(CASE WHEN ic.ORDER_TYPE=701 THEN 1 WHEN ic.ORDER_TYPE=702 THEN 2 END) AS 'OPERATE_TYPE'
			FROM cn_inventory_count_detail cd LEFT JOIN cn_inventory_count ic ON cd.INVENTORY_COUNT_ID=ic.id
			WHERE  ic.UPDATED_TS < now() and ic.STATUS = 2
			AND (ic.REMARK IS NULL OR ic.REMARK NOT LIKE '%charles%')
		) pdl
		GROUP BY pdl.ITEM_ID,pdl.STORE_CODE,pdl.INVENTORY_TYPE
	) as pandian ON come.STORE_CODE = pandian.STORE_CODE AND come.ITEM_ID = pandian.ITEM_ID AND come.INVENTORYTYPE = pandian.INVENTORY_TYPE
where come.INVENTORYTYPE = 1);
"""
    jerry_cursor.execute(sqlcommand)
    
def getInfo(user,upc):
    global problemDict
    result = {}
    sqlcommand =""
    if user in cainiaoList:
        realUser=user
        sqlcommand = """SELECT * from table1 
where username = '{}' 
and upc like "%{}%"
order by qty desc;""".format(realUser,upc)
    elif user in jerryList:
        realUser=ECWMSJERRYDict[user]
        sqlcommand = """select ski.ID,ski.LENGTH,ski.WIDTH,ski.HEIGHT,ski.SKU_NAME_C,cb.USERNAME,ski.UPC,sb.STOCK from customer_basic cb
	left join storage_sku_info ski  
on cb.id = ski.CUSTOMER_ID
left join stock_basic sb on sb.CUSTOMER_ID = cb.id and sb.SKU_ID = ski.ID
where cb.username = '{}'
and ski.upc like "%{}%"
order by sb.STOCK desc;""".format(realUser,upc)
    
    if len(sqlcommand)!=0:
        jerry_cursor.execute(sqlcommand)
        rows = list(jerry_cursor)
        if len(rows) == 1:
            for row in rows:
                result["sku"] = row[0]
                result["length"] = row[1]
                result["width"] = row[2]
                result["height"] = row[3]
        elif len(rows)>=1:
            result["sku"] = rows[0][0]
            result["length"] = rows[0][1]
            result["width"] = rows[0][2]
            result["height"] = rows[0][3]
            
            dictString = user+upc
            if dictString not in problemDict:
                temp=[]
                temp.append("一对多")
                temp.append(user)
                temp.append(upc)
                temp.append(realUser)
                for row in rows:
                    temp.append(row[0])#sku
                    temp.append(row[6])#upc
                    temp.append(row[7])#qty
                
                
                problemDict[dictString]=temp
            
        elif len(rows)==0:
            dictString = user+upc
            if dictString not in problemDict:
                temp=[]
                temp.append("不存在")
                temp.append(user)
                temp.append(upc)
                temp.append(realUser)
                problemDict[dictString]=temp
                
            
        
    if user == "BUBL":
        result["sku"] = upc
        result["length"] = 0.01
        result["width"] = 0.01
        result["height"] = 0.01
    elif user == "LAMO":
        result["sku"] = upc
        result["length"] = 0.23
        result["width"] = 0.19
        result["height"] = 0.04
    
    if "sku" not in result:
        result["sku"] = " "
    if "length" not in result:
        result["length"] = " "
    if "width" not in result:
        result["width"] = " "
    if "height" not in result:
        result["height"] = " "
    
    return result
    
def getInfo_old(user,upc):
    result = {}
    sqlcommand =""
    if user in cainiaoList:
        realUser=user
        sqlcommand = """SELECT cmi.ITEM_ID,cmi.LENGTH,cmi.WIDTH,cmi.HEIGHT FROM `cn_user` cu 
left join cn_commodity_info cmi on  
cu.USER_ID = cmi.OWNER_USER_ID
where cu.SUBSCRIBER_NICK = '{}' 
and cmi.BAR_CODE like "%{}%";""".format(realUser,upc)
    elif user in jerryList:
        realUser=ECWMSJERRYDict[user]
        sqlcommand = """select ski.ID,ski.LENGTH,ski.WIDTH,ski.LENGTH from customer_basic cb
	left join storage_sku_info ski  
on cb.id = ski.CUSTOMER_ID
where username = '{}'
and ski.upc like "%{}%";""".format(realUser,upc)
    
    if len(sqlcommand)!=0:
        jerry_cursor.execute(sqlcommand)
        rows = list(jerry_cursor)
        if len(rows) == 1:
            for row in rows:
                result["sku"] = row[0]
                result["length"] = row[1]
                result["width"] = row[2]
                result["height"] = row[3]
        elif len(rows)>=1:
            temp=[]
            for row in rows:
                temp.append(row[0])
            if user in cainiaoList:
                problemDict["菜鸟\t"+user+"\t"+upc+"\t"+",".join(temp)] = "一对多"
            elif user in jerryList:
                problemDict["JERRY\t"+user+"\t"+upc+"\t"+",".join(temp)] = "一对多"                
        elif len(rows)==0:
            if user in cainiaoList:
                problemDict["菜鸟\t"+user+"\t"+upc] = "不存在"
            elif user in jerryList:
                problemDict["JERRY\t"+user+"\t"+upc] = "不存在" 
        
    if user == "BUBL":
        result["sku"] = upc
        result["length"] = 0.01
        result["width"] = 0.01
        result["height"] = 0.01
    elif user == "LAMO":
        result["sku"] = upc
        result["length"] = 0.23
        result["width"] = 0.19
        result["height"] = 0.04
    
    if "sku" not in result:
        result["sku"] = " "
    if "length" not in result:
        result["length"] = " "
    if "width" not in result:
        result["width"] = " "
    if "height" not in result:
        result["height"] = " "
    
    return result
        
def getLineNo(lineNo):
    return "{0:05d}".format(lineNo)
        
def connect_to_JERRY():
    cf = configparser.ConfigParser()
    cf.read(sqlConfigFile)
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
    
def nowDate():
    return datetime.now().strftime('%Y-%m-%d')
def nowDatenoSpace():
    return datetime.now().strftime('%Y%m%d')

def toUpShelfExcelRow(upFile):
    resultRows=[]
    for value in upFile:
        tmpRow=[]
        tmpRow.append("")
        tmpRow.append(value.user)
        tmpRow.append(value.outerNo)
        tmpRow.append(value.upc)
        tmpRow.append(value.sku)
        tmpRow.append(value.lineNo)
        tmpRow.append(value.chinese_name)
        tmpRow.append("A品")
        tmpRow.append(value.quantity)
        tmpRow.append("澳大利亚-悉尼")
        tmpRow.append(value.area)
        tmpRow.append(value.location)
        tmpRow.append(value.batchNo)
        tmpRow.append(value.proxyNo)
        tmpRow.append("")
        tmpRow.append("")
        tmpRow.append("")
        tmpRow.append("")
        tmpRow.append(value.expireDate)
        tmpRow.append(nowDate())
        tmpRow.append("chriswang")
        
        resultRows.append(tmpRow)
    return resultRows
    
def toASNExcelRow(asnFile):
    resultRows=[]
    for key in asnFile:
        tmpRow=[]
        value = asnFile[key]
        tmpRow.append(value.outerNo)
        tmpRow.append(value.sku)
        tmpRow.append(value.chinese_name)
        tmpRow.append(value.upc)
        tmpRow.append(value.quantity)
        tmpRow.append("")
        tmpRow.append("")
        tmpRow.append(value.batchNo)
        tmpRow.append(value.proxyNo)
        tmpRow.append(value.length)
        tmpRow.append(value.width)
        tmpRow.append(value.height)
        tmpRow.append(value.expireDate)
        
        resultRows.append(tmpRow)
    return resultRows
    
def writeAsnFiles(asnFileList,user):
    output_userdir_ASN=output_dir_ASN+"\\"+ECWMS_to_RUNBOW_Dict[user]  
    if not os.path.isdir(output_userdir_ASN):
        os.mkdir(output_userdir_ASN) 
    
    n = 1
    for asnFile in asnFileList:
        fileName = output_userdir_ASN+"\\"+"ASN-"+ECWMS_to_RUNBOW_Dict[user]+"-"+str(n)+".xls"
        
        #如果存在同名文件则删除
        if os.path.isfile(fileName):
            os.remove(fileName)
        
        outputrows = toASNExcelRow(asnFile)
        
        file = xlwt.Workbook('utf-8')
        table = file.add_sheet("预入库单主信息")
        k=0
        for cell in ASN_First_page_first_line:
            table.write(0,k,cell)
            k+=1
            
        ASN_First_page_second_line=[outputrows[0][0],"采购入库",nowDate(),ECWMS_to_RUNBOW_Dict[user],"澳大利亚-悉尼","转仓测试"]
        k=0
        for cell in ASN_First_page_second_line:
            table.write(1,k,cell)
            k+=1
        table1=file.add_sheet("预入库单明细信息")
        k=0
        for cell in ASN_Second_page_first_line:
            table1.write(0,k,cell)
            k+=1
        
        m=1
        for row in outputrows:
            k=0
            for cell in row:
                if isinstance(cell,float):
                    table1.write(m,k,('%f' % cell).rstrip('0').rstrip('.'))
                elif cell is not None:
                    if k>=9 and k<=11 and cell is not None and not isinstance(cell,str):
                        table1.write(m,k,('%f' % cell).rstrip('0').rstrip('.'))
                    else:
                        table1.write(m,k,str(cell))
                k+=1
            m+=1

        file.save(fileName)
        n+=1

def writeUpshelfFiles(upFileList,user):
    output_userdir_upshelf = output_dir_Located +"\\"+ ECWMS_to_RUNBOW_Dict[user]
    if not os.path.isdir(output_dir_Located):
        os.mkdir(output_dir_Located) 
    
    if not os.path.isdir(output_userdir_upshelf):
        os.mkdir(output_userdir_upshelf) 
    n = 1
    for upFile in upFileList:
        fileName = output_userdir_upshelf+"\\"+"UpShelf-"+ECWMS_to_RUNBOW_Dict[user]+"-"+str(n)+".xls"
        
        #如果存在同名文件则删除
        if os.path.isfile(fileName):
            os.remove(fileName)
        
        outputrows = toUpShelfExcelRow(upFile)
        

        file = xlwt.Workbook('utf-8')
        table = file.add_sheet("上架单信息")
        k=0
        for cell in UpShelf_FirstLine:
            table.write(0,k,cell)
            k+=1
        
        m=1
        for row in outputrows:
            k=0
            for cell in row:
                if isinstance(cell,float):
                    table.write(m,k,('%f' % cell).rstrip('0').rstrip('.'))  
                elif cell is not None:
                    table.write(m,k,str(cell))
                k+=1
            m+=1

        file.save(fileName)
        n+=1

    
def getstr(rowNumber,str,sh):
    return sh.cell_value(rowNumber,ecwmsDict.index(str))

def readLine(rowNumber,sh):
    lineData=[]
    if getstr(rowNumber,"LotATT04",sh) == "Metcash" or getstr(rowNumber,"LotATT04",sh) == "EWES":
        return None
    lineData.append(getstr(rowNumber,"LotATT04",sh))
    lineData.append(getstr(rowNumber,"UPC",sh))
    lineData.append(getstr(rowNumber,"Descr C",sh))
    lineData.append(getstr(rowNumber,"Descr E",sh))
    lineData.append(getstr(rowNumber,"LOC",sh))
    lineData.append(getstr(rowNumber,"InvQty",sh))
    if(getstr(rowNumber,"Expiration Date",sh).startswith("999")):
        lineData.append("")
    else:
        lineData.append(getstr(rowNumber,"Expiration Date",sh))
    lineData.append(getstr(rowNumber,"Inbount Date",sh))
    return ECWMS(lineData)

# add
def ASN_add(asnRow,quantity):
    asnRow.quantity = asnRow.quantity+quantity
    return asnRow
    
# from ECWMS to ASN
def to_ASN(row,fileSecquence,lineNo):
    lineData=[]
    # print(row.user+row.upc)
    infoDict = getInfo(row.user,row.upc)
    lineData.append("ASN"+ECWMS_to_RUNBOW_Dict[row.user]+nowDatenoSpace()+"-"+str(fileSecquence))
    lineData.append(infoDict["sku"])
    lineData.append(row.upc)
    lineData.append(row.quantity)
    #batchNo
    lineData.append(("B"+row.inboundDate+"_"+row.expireDate).replace("-",""))
    lineData.append("")
    lineData.append(infoDict["length"])
    lineData.append(infoDict["width"])
    lineData.append(infoDict["height"])
    lineData.append(row.expireDate)
    lineData.append(ECWMS_to_RUNBOW_Dict[row.user])
    lineData.append(row.chinese_name.replace(",",""))
    lineData.append(lineNo)
    return ASN(lineData)

# to upShelf
def to_UPSHELF(userRows,file):
    singleUpfile=[]
    for key in file:
        tmpASNdata = file[key]
        tmplineData = UPSHELF()
        tmplineData.user = tmpASNdata.user
        tmplineData.outerNo = tmpASNdata.outerNo
        tmplineData.upc=tmpASNdata.upc
        tmplineData.sku=tmpASNdata.sku
        tmplineData.lineNo=getLineNo(tmpASNdata.lineNo)
        tmplineData.chinese_name=tmpASNdata.chinese_name
        tmplineData.batchNo=tmpASNdata.batchNo
        tmplineData.proxyNo=" "
        for row in userRows:
            # print(key+"\t"+row.upc)
            if(key == row.upc and tmpASNdata.expireDate == row.expireDate):
                outputLineData = copy.deepcopy(tmplineData)
                outputLineData.quantity=row.quantity
                #outputLineData.location=row.location
                outputLineData.area,outputLineData.location=checkArea(row.location)
                outputLineData.expireDate=row.expireDate
                outputLineData.inboundDate=nowDate()
                singleUpfile.append(outputLineData)
    
    return singleUpfile
        
    
class UPSHELF:
    # def __init__(self):
    
    # def __init__(self,row):
        # self.user = row[0]
        # self.outerNo = row[1]
        # self.upc=row[2]
        # self.sku=row[3]
        # self.lineNo=row[4]
        # self.chinese_name=row[5]
        # self.quantity=row[6]
        # self.area=row[7]
        # self.location=row[8]
        # self.batchNo=row[9]
        # self.proxyNo=row[10]
        # self.expireDate=row[11]
        # self.createDate=row[12]
        
    def toJSON(self):
        return simplejson.dumps(self, default=lambda obj: obj.__dict__,use_decimal=True) 
    
class ECWMS:
    def __init__(self,row):
        self.user = row[0]
        self.upc = row[1]
        self.chinese_name=row[2]
        self.english_name=row[3]
        self.location=row[4]
        self.quantity=row[5]
        self.expireDate=row[6]
        self.inboundDate=row[7]
        
    def toJSON(self):
        return simplejson.dumps(self, default=lambda obj: obj.__dict__,use_decimal=True) 

class ASN:
    def __init__(self,row):
        self.outerNo = row[0]
        self.sku = row[1]
        self.upc=row[2]
        self.quantity=row[3]
        self.batchNo=row[4]
        self.proxyNo=row[5]
        self.length=row[6]
        self.width=row[7]
        self.height=row[8]
        self.expireDate=row[9]
        self.user=row[10]
        self.chinese_name=row[11]
        self.lineNo=row[12]
        
    def toJSON(self):
        return simplejson.dumps(self, default=lambda obj: obj.__dict__,use_decimal=True)        

def getUserRows(user,originRows):
    userECWMSRows = []
    for row in originRows:
        if user == row.user:
            userECWMSRows.append(row)
    return sorted(userECWMSRows, key=operator.attrgetter('upc'))

def readFile(filename):
    global users
    print("开始转换 "+filename)
    inputfile = input_dir + "\\" + filename
    book = xlrd.open_workbook(inputfile)
    sh = book.sheet_by_index(0)
    originRows=[]
    for rowNumber in range(sh.nrows):
    #for rowNumber in tqdm(range(sh.nrows)):
        if rowNumber != 0:
            oneLine = readLine(rowNumber,sh)
            if oneLine is not None:
                originRows.append(oneLine)
            
    for row in originRows:
        if row.user not in users:
            users.append(row.user)
    
    for user in users:
        userECWMSRows = getUserRows(user,originRows)
        createASNFile(userECWMSRows,user)
def writeProblemList():
    print("Start outputing problem file")
    fileName = "problemList.xlsx"    
    #如果存在同名文件则删除
    if os.path.isfile(fileName):
        os.remove(fileName)
    
    file = xlsxwriter.Workbook(fileName)
    table = file.add_worksheet('问题件')
    table.write_row('A1',problem_first_line)
    n=2
    for key in problemDict:
        table.write_row('A'+str(n),problemDict[key])
        n+=1
    file.close()
    print("Problem file finished")
    
    

        
def createASNFile(userECWMSRows,user):
    print("Start outputing "+user+" file")
    if not os.path.isdir(output_dir_ASN):
        os.mkdir(output_dir_ASN)
  
    
    userRows = copy.deepcopy(userECWMSRows)
    asnFileList=[]
    
    # 转换asn
    # 文件名顺序
    fileSecquence=1    
    while len(userRows)!=0:
        tmpFile = {}
        #负责pop
        i = 0
        #行数顺序
        lineNo = 1
        while i < len(userRows):
            if userRows[i].upc not in tmpFile:
                asnLine = to_ASN(userRows[i],fileSecquence,lineNo)
                #这里决定是否要导 sku为空的
                if asnLine.sku!=" ": 
                    tmpFile[userRows[i].upc] = asnLine
                    lineNo+=1
                userRows.pop(i)
                i-=1
            elif userRows[i].upc in tmpFile and userRows[i].expireDate == tmpFile[userRows[i].upc].expireDate:
                tmpFile[userRows[i].upc] = ASN_add(tmpFile[userRows[i].upc],userRows[i].quantity)
                userRows.pop(i)
                i-=1
            
            i+=1
        
        asnFileList.append(tmpFile)
        fileSecquence+=1
    #转换上架单
    upFileList=[]
    for file in asnFileList:
        up_tmp_file = to_UPSHELF(userECWMSRows,file)
        upFileList.append(up_tmp_file)
    
    
    #打印asn
    writeAsnFiles(asnFileList,user)
    #打印上架单
    writeUpshelfFiles(upFileList,user)
            
    
if __name__ == '__main1__':
    lineNo = 5
    print(getLineNo(lineNo))
    

if __name__ == '__main__':
    jerry = connect_to_JERRY()
    jerry_cursor = jerry.cursor()
    initial()
    print("导出数据转化所需数据开始！")
    try:
        files = os.listdir(input_dir)
    except FileNotFoundError:
        print("请把待转换的文件放到程序的"+input_dir+"、目录下！")
        ord(msvcrt.getch())
        quit()
    
    n=0
    for file in files:
        if file.endswith(".xlsx") and not file.startswith("~$"):
            readFile(file)
            n+=1
    
    #打印问题单
    writeProblemList()
    
    if n==0:
        print("未发现扩展名为xlsx的文件，按D键退出")
    else:
        print("共计"+str(n)+"个文件，按D键退出")
    
    jerry_cursor.close()
    jerry.close()
    #打印最后数据

    while True:
        if ord(msvcrt.getch()) in [68, 100]:
            break