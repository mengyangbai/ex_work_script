# !/usr/bin/env python3
# @authoer bain.bai
# 计算菜鸟仓储费之用
# 月末开始，计算本月 从2017-05-31 开始 从2017-05-01 到2017-05-31
# 菜鸟库存和流水规律，库存和流水库存在cn_account_snapshot_detail表
# 流水在cn_account_snapshot_detail 表
# 每天北京时间0点，生成昨天一天的数据，而我们系统时标准时间，所以是2017-05-01 16：00 生成的2017-05-01 一天的数据
# 生成时间有create date来判定

import configparser
import pypyodbc
import pymysql
import sys,os
import xlsxwriter
import locale

from tqdm import tqdm
from datetime import datetime, timedelta

def connect_to_JERRY():
    cf = configparser.ConfigParser()
    cf.read('..\\config\\sql.ini')
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
    
#抓取数据库数据输出到菜鸟
def exportToCaiNiao(start_time,end_time,delta_days):
    price_rate = 1.5
    newlocale = locale.setlocale(locale.LC_CTYPE, 'chinese')

    jerry = connect_to_JERRY()
    jerry_cursor = jerry.cursor()
    sqlcommand ="""create TEMPORARY table temp1 (SELECT casd.ITEM_ID, 
        sum(casd.QUANTITY) stockNum,DATE_FORMAT(cas.CREATED_TS,'%Y%m%d') time,
        casd.STORE_CODE FROM `cn_account_snapshot_detail` casd 
	left join cn_account_snapshot cas on cas.ID = casd.ACCOUNT_SNAPSHOT_ID
	where cas.CREATED_TS between '"""+start_time + "' and '"+end_time+"' group by time,casd.ITEM_ID);"
    jerry_cursor.execute(sqlcommand)
    sqlcommand="""create TEMPORARY table  temp2 (select cwcd.ITEM_ID,sum(cwcd.QUANTITY) sellNum,DATE_FORMAT(cwcd.CREATED_TS,'%Y%m%d') time from cn_water_check_detail cwcd 
	where cwcd.TABLE_NAME = 'cn_consign_order_notify' and cwcd.CREATED_TS between '"""+start_time +"' and '"+end_time +"' group by time,cwcd.ITEM_ID);"
    jerry_cursor.execute(sqlcommand)
    
    print("Data warehousing processing")
    
    sqlcommand="alter table temp1 convert to character set utf8 collate utf8_unicode_ci;"
    jerry_cursor.execute(sqlcommand)
    sqlcommand="alter table temp2 convert to character set utf8 collate utf8_unicode_ci;"
    jerry_cursor.execute(sqlcommand)
    sqlcommand="""select temp1.STORE_CODE,temp1.time,
                    cu.SUBSCRIBER_NICK,temp1.ITEM_ID,
                    cci.`NAME`,cci.LENGTH*100,cci.WIDTH*100,cci.HEIGHT*100,
                    temp1.stockNum,temp2.sellNum
                    from temp1 left join temp2 on temp1.ITEM_ID = temp2.ITEM_ID and temp1.time=temp2.time
                    left join cn_commodity_info cci on temp1.ITEM_ID = cci.ITEM_ID
                    left join cn_user cu on cci.OWNER_USER_ID = cu.USER_ID
                    where 
                    temp1.stockNum !=0
                    order by temp1.time asc,cu.SUBSCRIBER_NICK asc;"""
                    
    # sqlcommand="""select temp1.STORE_CODE,temp1.time,
                    # cu.SUBSCRIBER_NICK,temp1.ITEM_ID,
                    # cci.`NAME`,cci.LENGTH*100,cci.WIDTH*100,cci.HEIGHT*100,
                    # temp1.stockNum,temp2.sellNum
                    # from temp1 left join temp2 on temp1.ITEM_ID = temp2.ITEM_ID and temp1.time=temp2.time
                    # left join cn_commodity_info cci on temp1.ITEM_ID = cci.ITEM_ID
                    # left join cn_user cu on cci.OWNER_USER_ID = cu.USER_ID
                    # where 
                    # temp1.stockNum !=0 and 
                    # cu.SUBSCRIBER_NICK !='chemistwarehouse海外旗舰'
                    # order by temp1.time asc,cu.SUBSCRIBER_NICK asc;"""
    jerry_cursor.execute(sqlcommand)
    
    
    #如果不存在仓储费目录则创建
    if not os.path.isdir("仓储费"):
        os.mkdir("仓储费")
        
    filename = "仓储费\\菜鸟仓储费"+".xlsx"
    
    #如果存在同名文件则删除
    if os.path.isfile(filename):
        os.remove(filename)
        
    
    start_date  = datetime.strptime(start_time,'%Y-%m-%d %H:%M')
    
    month=start_date.strftime('%#m月')
    year_month=start_date.strftime('%Y年%#m月')
    
    #把结果录入result.xls
    file = xlsxwriter.Workbook(filename)
    table1 = file.add_worksheet('汇总')
    table2 = file.add_worksheet(month)
    
    #设置宽度
    table1.set_column(0,18,13)
    table2.set_column(0,18,10)
    
    print("Start writing to "+filename)
    
    table1.write(6,2,'GFC仓储费')
    table1.write(7,2,'账期')
    table1.write(7,3,'仓储费$')
    table1.write(7,4,'备注')
    table1.write(8,2,year_month)
    table1.write(8,4,'未结算')
    
    table2.write(0,0,'仓库Code')
    table2.write(0,1,'业务日期')
    table2.write(0,2,'卖家nick')
    table2.write(0,3,'后端货品id')
    table2.write(0,4,'后端货品名称')
    table2.write(0,5,'长(cm)')
    table2.write(0,6,'宽(cm)')
    table2.write(0,7,'高(cm)')
    table2.write(0,8,'总库存件')
    table2.write(0,9,'销售出库')
    table2.write(0,10,'仓库')
    table2.write(0,11,'商品体积(立方厘米)')
    table2.write(0,12,'sku总体积(立方米)')
    table2.write(0,13,'行标签')
    table2.write(1,13,'EWE澳洲仓库')
    table2.write(2,13,'总计')
    table2.write(5,13,'CP端')
    table2.write(6,13,'仓库')
    table2.write(7,13,'单价（￥元/立方/天）')
    table2.write(8,13,'当月日均结余库存件数')
    table2.write(9,13,'当月日均销售件数')
    table2.write(10,13,'周转天数')
    table2.write(11,13,'计费天数')
    table2.write(12,13,'日均SKU总体积')
    table2.write(13,13,'总计')
    table2.write(6,14,'EWE仓')
    table2.write(7,14,price_rate)
    table2.write(8,15,'除以自然日')
    table2.write(9,15,'除以自然日')
    table2.write(0,14,'求和项：总库存件数')
    table2.write(0,15,'求和项：销售出库件数')
    table2.write(0,16,'求和项：sku总体积')
    
    n=1
    for row in tqdm(jerry_cursor):
        table2.write(n,0,row[0])
        table2.write(n,1,row[1])
        table2.write(n,2,row[2])
        table2.write(n,3,row[3])
        table2.write(n,4,row[4])
        table2.write(n,5,row[5])
        table2.write(n,6,row[6])
        table2.write(n,7,row[7])
        table2.write(n,8,row[8])
        if row[9] is not None:
            table2.write(n,9,row[9])
        else:
            table2.write(n,9,0)
        table2.write(n,10,'EWE仓库')
        volume = "=F"+str(n+1)+"*G"+str(n+1)+"*H"+str(n+1)
        table2.write(n,11,volume)
        skuvolume="=L"+str(n+1)+"/1000000*I"+str(n+1)
        table2.write(n,12,skuvolume)
        n+=1
    
    #最后的输出
    total_number_stock="=SUM(I2:I"+str(n)+")"
    total_number_sell="=SUM(J2:J"+str(n)+")"
    total_volume="=SUM(M2:M"+str(n)+")"
    
    
    table2.write(1,14,total_number_stock)
    table2.write(1,15,total_number_sell)
    table2.write(1,16,total_volume)
    
    table2.write(2,14,total_number_stock)
    table2.write(2,15,total_number_sell)
    table2.write(2,16,total_volume)
    
    day_of_stock ="=O3/"+ str(delta_days)
    day_of_sell ="=P3/"+str(delta_days)
    estimated_days="=O9/O10"
    calculated_days="=IF((O11-90)>30,30,IF((O11-90)<0,0,(O11-90)))"
    day_of_volume ="=Q3/"+str(delta_days)
    result="=O8*O12*O13"
    
    table2.write(8,14,day_of_stock)
    table2.write(9,14,day_of_sell)
    table2.write(10,14,estimated_days)
    table2.write(11,14,calculated_days)
    table2.write(12,14,day_of_volume)
    table2.write(13,14,result)
    
    #关闭file
    file.close()



if __name__ == '__main__':
    # end_time =datetime.now().strftime('%Y-%m-%d')    
    # start_time =(datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')
    end_time='2018-03-01 15:00'
    start_time='2018-02-01 15:00'
    delta_days=28
    exportToCaiNiao(start_time,end_time,delta_days)
        