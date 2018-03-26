# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author 白孟阳
# 检查库区库位是否合法
import os
import xlrd
import functools

_cainiao_file = "ewe菜鸟仓储费.xlsx"
_ewe_file = "菜鸟仓储费.xlsx"
cainiao_set = set()
ewe_set=set()

def log(func):
    @functools.wraps(func)
    def wrapper(*args, **kw):
        print("{} start".format(func.__name__))
        return func(*args, **kw)
    return wrapper

@log
def initial():
    with xlrd.open_workbook(_cainiao_file) as f:
        sh = f.sheet_by_index(0)
        for row_number in range(sh.nrows):
            if row_number != 0:
                cainiao_set.add((sh.cell_value(row_number,1),sh.cell_value(row_number,10)/10,sh.cell_value(row_number,11)/10,sh.cell_value(row_number,12)/10,sh.cell_value(row_number,17)))
    
    with xlrd.open_workbook(_ewe_file) as f:
        sh = f.sheet_by_index(1)
        for row_number in range(sh.nrows):
            if row_number != 0:
                ewe_set.add((sh.cell_value(row_number,3),sh.cell_value(row_number,5),sh.cell_value(row_number,6),sh.cell_value(row_number,7),sh.cell_value(row_number,11)*1000))
    
def getFirst(set):
    return set[0]
    
if __name__ == '__main__':
    initial()
    cainiao_rest_set = cainiao_set-ewe_set
    ewe_rest_set = ewe_set - cainiao_set
    
    
    print("菜鸟不同")
    
    for set in sorted(cainiao_rest_set,key=getFirst):
        print(set)
    
    print("EWE不同")
    for set in sorted(ewe_rest_set,key=getFirst):
        print(set)
