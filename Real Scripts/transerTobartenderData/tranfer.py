# !/usr/bin/env python3
# @Author 白孟阳
# 转化导出表到面单可读数据的脚本
import xlrd
import xlsxwriter
import msvcrt
import os
from tqdm import tqdm

# Basic setting
input_dir = '报表'
output_dir = '面单'
inputdict=['序号','分运单号','货物品名','件数','重量KG','数量','单位','货币编码','价格','个人完税税号','型号','国别代码','原产国','HS编码','收件人ID','收件人','地址','收件人电话','TO','落地配单号','寄件人公司','寄件人','寄件人电话','FROM','货主城市','备注','英文品名','规格']
outputFirstLine=['EWE单号','转单号','收件人','地址','电话','核重','物品1','规格1','数量1','价值1','物品2','规格2','数量2','价值2','物品3','规格3','数量3','价值3','物品4','规格4','数量4','价值4']

def getstr(rowNumber,str,sh):
    return sh.cell_value(rowNumber,inputdict.index(str))

def readLine(rowNumber,sh,oneLine,outputrows):
    eweNo = getstr(rowNumber,"分运单号",sh)
    '''
    先不考虑为空的合并状况，
    如果合并，就把新的加到旧的后面去
    需额外判断几个字段，否则抛错误
    '''
    if rowNumber == 1:
        tmpLine =[]
        tmpLine.append(eweNo)
        tmpLine.append(getstr(rowNumber,"落地配单号",sh))
        tmpLine.append(getstr(rowNumber,"收件人",sh))
        tmpLine.append(getstr(rowNumber,"地址",sh))
        tmpLine.append(getstr(rowNumber,"收件人电话",sh))
        tmpLine.append(getstr(rowNumber,"重量KG",sh))
        tmpLine.append(getstr(rowNumber,"货物品名",sh))
        tmpLine.append(getstr(rowNumber,"规格",sh))
        tmpLine.append(getstr(rowNumber,"数量",sh))
        tmpLine.append(getstr(rowNumber,"价格",sh))
        return tmpLine
    elif len(oneLine) !=0 and (oneLine[0] == eweNo or eweNo ==''):
        tmpLine=oneLine
        tmpLine.append(getstr(rowNumber,"货物品名",sh))
        tmpLine.append(getstr(rowNumber,"规格",sh))
        tmpLine.append(getstr(rowNumber,"数量",sh))
        tmpLine.append(getstr(rowNumber,"价格",sh))
        outputrows.pop()
        return tmpLine
    else:
        tmpLine =[]
        tmpLine.append(eweNo)
        tmpLine.append(getstr(rowNumber,"落地配单号",sh))
        tmpLine.append(getstr(rowNumber,"收件人",sh))
        tmpLine.append(getstr(rowNumber,"地址",sh))
        tmpLine.append(getstr(rowNumber,"收件人电话",sh))
        tmpLine.append(getstr(rowNumber,"重量KG",sh))
        tmpLine.append(getstr(rowNumber,"货物品名",sh))
        tmpLine.append(getstr(rowNumber,"规格",sh))
        tmpLine.append(getstr(rowNumber,"数量",sh))
        tmpLine.append(getstr(rowNumber,"价格",sh))
        return tmpLine
        


def transferFile(filename):    
    print("开始转换 "+filename)
    inputfile = input_dir + "\\" + filename
    book = xlrd.open_workbook(inputfile)
    sh = book.sheet_by_index(0)
    outputrows=[]
    oneLine=[]
    #for rowNumber in range(sh.nrows):
    for rowNumber in tqdm(range(sh.nrows)):
        if rowNumber != 0:
            oneLine = readLine(rowNumber,sh,oneLine,outputrows)
            outputrows.append(oneLine)
            
    writeFile(outputrows,filename)
        
def writeFile(outputrows,filename):
    print(filename+" 开始输入")
    outputfile = output_dir+"\\"+filename
    #如果不存在妥投目录则创建
    if not os.path.isdir(output_dir):
        os.mkdir(output_dir)
    
    #如果存在同名文件则删除
    if os.path.isfile(outputfile):
        os.remove(outputfile)
    
    file = xlsxwriter.Workbook(outputfile)
    table = file.add_worksheet()
    
    table.write_row('A1',outputFirstLine)
    n=1
    for row in outputrows:
        table.write_row('A'+str(n),row)
        n+=1
    
    file.close()
    print(filename+" 输出完成")



if __name__ == '__main__':
    print("导出数据转化面单所需数据开始！")
    try:
        files = os.listdir(input_dir)
    except FileNotFoundError:
        print("请把待转换的文件放到程序的"+input_dir+"、目录下！")
        ord(msvcrt.getch())
        quit()
    
    n=0
    for file in files:
        if file.endswith(".xlsx") and not file.startswith("~$"):
            transferFile(file)
            n+=1
    
    if n==0:
        print("未发现扩展名为xlsx的文件，按D键退出")
    else:
        print("共计"+str(n)+"个文件，按D键退出")
    while True:
        if ord(msvcrt.getch()) in [68, 100]:
            break