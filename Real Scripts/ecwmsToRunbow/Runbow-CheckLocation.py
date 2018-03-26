# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author 白孟阳
# Check runbow 
import xlrd
import xlsxwriter
import os

input_dir = 'ECWMS'
ecwmsDict = ["LotATT04","UPC","Descr C","Descr E","LOC","ID","InvQty","AllQty","HoldQty","AvblQty","Production Date","Expiration Date","Inbount Date"]

def getstr(rowNumber,str,sh):
    return sh.cell_value(rowNumber,ecwmsDict.index(str))
    
def readLine(rowNumber,sh):
    lineData=[]
    if getstr(rowNumber,"LOC",sh) == "DamageItems" or getstr(rowNumber,"LOC",sh) == "expireItems" or getstr(rowNumber,"LOC",sh) == "ExpiredItems" or getstr(rowNumber,"LotATT04",sh) == "Metcash" or getstr(rowNumber,"LotATT04",sh) == "EWES":
        return None
    lineData.append(getstr(rowNumber,"LotATT04",sh))
    lineData.append(getstr(rowNumber,"UPC",sh))
    lineData.append(getstr(rowNumber,"LOC",sh))
    lineData.append(getstr(rowNumber,"InvQty",sh))
    if(getstr(rowNumber,"Expiration Date",sh).startswith("999")):
        lineData.append("")
    else:
        lineData.append(getstr(rowNumber,"Expiration Date",sh))
    return lineData

    
'''
    考虑效期
    au1为爆品
    Chullora-storage为非爆品
    日期格式为YYYY-MM-DD
    爆品比较爆品  效期1和效期2相比，时间1《时间2 返回 1 时间1》=时间2 返回2
    爆品比较非爆品 效期1和效期2相比，时间1《=时间2 返回 2 时间1>时间2 返回0
    非爆品比较爆品 效期1和效期2相比，时间1《时间2 返回 0 时间1>=时间2 返回1
    非爆品比较非爆品 效期1和效期2相比，时间1《时间2 返回 1 时间1>=时间2 返回2
'''
def compare(data1,data2):
    def is_best_seller(area):
        if area == "AU1":
            return True
        elif area == "Chullora-storage":
            return False
        raise Exception("what the fuck")
        
    '''
        1>2 return 1
        1=2 return 0
        1<2 rerurn -1
    '''
    def compare_time(time1,time2):
        number1 = 0
        number2 = 0
        if time1 == "":
            number1 = 999999999
        else:
            row = time1.split("-")
            number1 =int(row[0])*365+int(row[1].lstrip("0"))*30+ int(row[2].lstrip("0"))
        if time2 == "":
            number2 = 999999999
        else:
            row = time2.split("-")
            number2 = int(row[0])*365+int(row[1].lstrip("0"))*30+ int(row[2].lstrip("0"))
        if number1==number2:
            return 0
        elif number1 > number2:
            return 1
        elif number1 < number2:
            return -1
    
        
    area1 = data1[0]
    area2 = data2[0]
    time1 = data1[3]
    time2 = data2[3]
    #print(area1+"||"+area2)
    #print(time1+"||"+time2)
    if is_best_seller(area1) and is_best_seller(area2):
        if compare_time(time1,time2) == -1:
            return 1
        elif compare_time(time1,time2) == 1 or compare_time(time1,time2) == 0:
            return 2
    elif is_best_seller(area1) and not is_best_seller(area2):
        if compare_time(time1,time2) == -1 or compare_time(time1,time2) == 0:
            return 2
        elif compare_time(time1,time2) == 1:
            return 0
    elif not is_best_seller(area1) and is_best_seller(area2):
        if compare_time(time1,time2) == -1:
            return 0
        elif compare_time(time1,time2) == 1 or compare_time(time1,time2) == 0:
            return 1
    elif not is_best_seller(area1) and not is_best_seller(area2):
        if compare_time(time1,time2) == -1:
            return 1
        elif compare_time(time1,time2) == 1 or compare_time(time1,time2) == 0:
            return 2
            
    raise Exception("what the fuck")

    
'''
    if data size = 0 just return false
    first transfer to new runbow location
    then compare
'''
def check(data):
    if len(data) == 1:
        return False
    else:
        current_latest=data[0]
        for row in data:
            if compare(current_latest,row)==0:
                return True
            elif compare(current_latest,row)==1:
                current_latest = row
                
        return False
        
    
def checkArea(str):
    result = ""
    
    rows = str.split("-")
    # if(len(rows)==4):
    target = rows[2]
    if target.isdigit():
        result = "AU1"
    else:
        result = "Chullora-storage"
        str="ST"+str
    # else:
        # print(str)
        # result = "注意"
    return result,str    
    
    
def data_transfer(originRows):
    result_data={}
    for row in originRows:
        key = row[0]+"||"+row[1]
        if key in result_data:
            area,location = checkArea(row[2])
            tmp = [area,location,row[3],row[4]]
            result_data[key].append(tmp)
        else:
            tmp_list = []
            area,location = checkArea(row[2])
            tmp = [area,location,row[3],row[4]]
            tmp_list.append(tmp)
            result_data[key] = (tmp_list)
            
    for key in result_data:
        if check(result_data[key]):
            pass
        else:
            result_data[key] = 0
            
    return result_data
    
    
def read_file(filename):
    inputfile = input_dir + "\\" + filename
    originRows=[]
    with xlrd.open_workbook(inputfile) as book:
        sh = book.sheet_by_index(0)
        for rowNumber in range(sh.nrows):
            if rowNumber != 0:
                oneLine = readLine(rowNumber,sh)
                if oneLine is not None:
                    originRows.append(oneLine)
                    
    return data_transfer(originRows)
        
def output_to_location_file(result):
    output_file = "location_report.xlsx"
    with xlsxwriter.Workbook(output_file) as file:
        table = file.add_worksheet('汇总')
        table.write_row('A1',["商家","UPC","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间","库区","库位","个数","过期时间"])
        n=2
        for key in result:
            if result[key] != 0:
                row = key.split("||")
                for value in result[key]:
                    row.extend(value)
                table.write_row('A'+str(n),row)
                n+=1
                                         
        

if __name__=='__main__':    
    try:
        files = os.listdir(input_dir)
    except FileNotFoundError:
        print("请把待转换的文件放到程序的"+input_dir+"、目录下！")
        ord(msvcrt.getch())
        quit()
    
    result={}
    n=0
    for file in files:
        if file.endswith(".xlsx") and not file.startswith("~$"):
            result = read_file(file)
            n+=1
            
    output_to_location_file(result)
            

    print("{} file has been checked, please check the location report".format(n))