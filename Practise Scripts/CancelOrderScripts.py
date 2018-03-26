##白取消用脚本
#cancel test
#记得要做失败以后的处理，不能这么草率了
import json
import configparser
import requests
import pymysql
from openpyxl import load_workbook
from tqdm import tqdm

def cancel_order_by_boxno(boxno=""):

    value ={}
    value['BoxNo']=boxno
    value['USERNAME']= USERNAME
    value['APIPASSWORD']=clientcf.get("AuSydHaituncun","APIPASSWORD")


    url = 'http://jerryapi.ewe.com.au/eweApi/ewe/api/cancelOrder'
    headers = {'Content-Type': 'application/javascript'}

    #print(json.dumps(value))
    r = requests.get(url, headers=headers,data=json.dumps(value))
    print(r.text)

def select_boxno_by_referenceno(referenceno):
    sqlcommand = """create TEMPORARY table if not EXISTS temp1(
                        referenceno VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""
    cur.execute(sqlcommand)
    
    sqlcommand = """insert into temp1 values %s;"""
    cur.execute(sqlcommand % referenceno)
    sqlcommand = """select ib.BOX_NO,temp1.referenceno from temp1 
                    left JOIN order_basic ob on temp1.referenceno = ob.REFERENCE_NO
                    LEFT JOIN inventory_basic ib on ob.ID = ib.ORDER_ID;"""
    
    cur.execute(sqlcommand)                
    for row in tqdm(cur):
        if row[0] is not None:
            cancel_order_by_boxno(row[0])
            #在这里最好做好处理下次
        else:
            nonelist.append(row[1])

if __name__ == '__main__':
    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')    
    clientcf = configparser.ConfigParser()
    clientcf.read('config\\client.ini')
    USERNAME = "AuSydHaituncun"
    #连接mysql
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()
    referenceno = []
    nonelist=[]
    wb = load_workbook('澳洲商户已取消订单3-1.xlsx')
    sheets = wb.get_sheet_names()
    for sheet in (sheets):
        ws = wb[sheet]
        for row in tqdm(ws.rows):
            if row[0].value != "Order Number":
                temp = row[0].value
                temp = temp.rstrip('\t')
                temp = temp.rstrip(' ')
                temp = "('"+temp+"')"
                referenceno.append(temp)
    
    select_boxno_by_referenceno(','.join(referenceno))
    
    #关闭mysql
    cur.close()
    conn.close()