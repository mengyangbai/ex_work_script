# !/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author Bain.Bai
# For implementing basic function of CCIC code scripts
import os,io,functools
from shutil import copyfile

dir_dict ={"墨尔本":"\\\\192.168.5.201\MelNormal\\04 物控部\口岸材料","悉尼":"\\\\192.168.5.214\Warehouse\ANDY新发货文件","布里斯班":"\\\\192.168.5.214\Warehouse\布里斯班发货文件"}
output_dir = 'ccic_report'

def get_and_check(airwaybill,city):
    if city in dir_dict:
        #Currently in city
        for root, dirs, files in os.walk(dir_dict[city]):
            for file in files:
                if(airwaybill in file and "Email Copy -" in file):
                    copyfile(os.path.join(root, file),os.path.join(output_dir,file))
                    print("Copying from {} to {}...".format(os.path.join(root, file),os.path.join(output_dir,file)))
                    return True,os.path.join(output_dir,file)
                    
        # for anotherCity in dir_dict:
            # if city!=anotherCity:
                # #Currently in city
                # for root, dirs, files in os.walk(dir_dict[anotherCity]):
                    # for file in files:
                        # if(airwaybill in file and "Email Copy -" in file):
                            # print("{} in wrong city, tell Haiyan, it should be {}".format(airwaybill,anotherCity))
                            # return False
                        
        print("{} currently not existed".format(airwaybill))
        return False
    else:
        for city in dir_dict:
            #Currently in city
            for root, dirs, files in os.walk(dir_dict[city]):
                for file in files:
                    if(airwaybill in file and "Email Copy -" in file):
                        print("{} in wrong city, tell Haiyan, it should be {}".format(airwaybill,city))
                        return False
                       
        print("Brisban: {}".format(airwaybill))
        return False


if __name__=='__main__':
    get_and_check("618-47362125","布里斯班")