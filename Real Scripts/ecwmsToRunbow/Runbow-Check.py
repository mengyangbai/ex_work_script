# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author 白孟阳
# 检查库区库位是否合法
import os
import xlrd
import functools
import xlsxwriter

_location_file = "Runbow 库位 (live).xls"
area_location_set = set()
output_file = "report.xlsx"

def log(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
        print("{} 初始化开始".format(func.__name__))
        return func(*args, **kw)
    return wrapper

@log
def initial():
    with xlrd.open_workbook(_location_file) as f:
        sh = f.sheet_by_index(0)
        for row_number in range(sh.nrows):
            if row_number != 0:
                area_location_set.add((sh.cell_value(row_number,0),sh.cell_value(row_number,1)))

if __name__ == '__main__':
    initial()
    problem_list=[]
    number_of_files,number_of_illegal_files=0,0
    
    for path, subdirs, files in os.walk("."):
        for name in files:
            if("RUNBOW-LOCATED" in path and name.endswith("xls")):
                number_of_files+=1
                with xlrd.open_workbook(os.path.join(path, name)) as f:
                    sh = f.sheet_by_index(0)
                    check_flag = False #避免多加
                    for row_number in range(sh.nrows):
                        if row_number != 0:
                            temp_variable = (sh.cell_value(row_number,10),sh.cell_value(row_number,11))
                            if temp_variable not in area_location_set:
                                if not check_flag:
                                    number_of_illegal_files+=1
                                check_flag=True
                                problem_list.append((path,name,row_number+1,sh.cell_value(row_number,10),sh.cell_value(row_number,11)))
    
    if os.path.isfile(output_file):
        os.remove(output_file)
    with xlsxwriter.Workbook(output_file) as file:
        table = file.add_worksheet('汇总')
        table.write_row('A1',["路径","文件名","行数","库区","库位"])
        n=2
        
        for row in problem_list:
            table.write_row('A'+str(n),row)
            n+=1
        
        table.write(n,0,"总共检查{}个文件，{}需核查".format(number_of_files,number_of_illegal_files))
    print("总共检查{}个文件，{}需核查".format(number_of_files,number_of_illegal_files))
    print("检查完毕")
    

