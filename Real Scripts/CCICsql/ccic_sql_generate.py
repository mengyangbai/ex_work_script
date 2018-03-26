# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author Bain.Bai
# For generate sql for CCIC code

import os
import xlrd
import copy

'''
    Some global setting
'''
START_DATE="2017-09-27 08:00"
EXPIRE_DATE="2018-09-27 08:00" #should be one year later than the start date
INPUT_DIR="CCIC"
list_ccic=[]
list_ccic_sequence=set()
list_ccic_code=set()
out_put_file="output"

'''
    For generate one single sql
'''
def export_sql(ccic_sequence,ccic_code):
    sql="INSERT INTO ccic_code (`CCIC_SEQUENCE`, `CCIC_CODE`, `BOX_NO`, `CREATE_DATE`, `LAST_MODIFIED_DATE`, `EXPIRED_DATE`) VALUES ('{}', '{}', null, '{}', '{}', '{}');".format(ccic_sequence, ccic_code, START_DATE, START_DATE, EXPIRE_DATE)
    return sql

def getstr(rowNumber,column_number,sh):
    return sh.cell_value(rowNumber,column_number)
    
def read_file(filename):
    print("Start reading file " + filename)
    inputfile = INPUT_DIR + "\\" + filename
    book = xlrd.open_workbook(inputfile)
    sh = book.sheet_by_index(0)
    for rowNumber in range(sh.nrows):
        if rowNumber != 0:
            ccic_sequence = str(getstr(rowNumber,1,sh))
            ccic_code = str(getstr(rowNumber,2,sh))
            if ccic_sequence not in list_ccic_sequence and ccic_code not in list_ccic_code:
                list_ccic_sequence.add(ccic_sequence)
                list_ccic_code.add(ccic_code)
                list_ccic.append((ccic_sequence,ccic_code))
            else:
                raise RuntimeError("there are same duplicate data")
    
 
if __name__=='__main__':
    try:
        files = os.listdir(INPUT_DIR)
    except FileNotFoundError:
        print("Please set the file under the "+INPUT_DIR+" directory")
        ord(msvcrt.getch())
        quit()
    
    for file in files:
        if file.endswith(".xls") and not file.startswith("~$"):
            read_file(file)
            
    try:
        files = os.listdir(".\\")
    except FileNotFoundError:
        print("Please set the file under the "+INPUT_DIR+" directory")
        ord(msvcrt.getch())
        quit()
        
    for file in files:
        if file.endswith(".sql"):
            os.remove(file)
     
    print("There are total "+str(len(list_ccic))+" code")
    
    output_list=[]
    
    tmp_list=[]
    while len(list_ccic) != 0:
        tmp_list.append(list_ccic.pop())
        if len(tmp_list)==20000:
            add_list=copy.deepcopy(tmp_list)
            output_list.append(add_list)
            tmp_list=[]
            
    n=1
    for output in output_list:
        with open(out_put_file+"~"+str(n)+".sql",'w') as f:
            for row in output:
                f.write(export_sql(row[0],row[1])+"\n")
        n+=1
            
    print("Work done!")