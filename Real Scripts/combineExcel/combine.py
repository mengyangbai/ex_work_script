# !/usr/bin/env python3
# @Author 白孟阳
# 把ECWMS的excel转换成runbow的两个excel
import xlrd
import os
import glob
import pandas as pd
import numpy as np

input_dir = 'input'

def read_file(file):
    pass
    
 
if __name__=='__main__':
        
    all_data = pd.DataFrame()
    

    n=0
    for file in glob.glob("input/*.xlsx"):
        df = pd.read_excel(file)
        all_data = all_data.append(df,ignore_index=True)
        
    all_data.describe()