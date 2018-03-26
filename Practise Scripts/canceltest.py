#cancel test
import json
import configparser
import requests
import pymysql

def cancel_order_by_boxno(boxno=""):

    value ={}
    value['BoxNo']=boxno
    value['USERNAME']= USERNAME
    value['APIPASSWORD']=clientcf.get("AuSydHaituncun","APIPASSWORD")


    url = 'http://localhost:8080/eweApi/ewe/api/cancelOrder'
    headers = {'Content-Type': 'application/javascript'}

    print(json.dumps(value))
    #r = requests.get(url, headers=headers,data=json.dumps(value))
    #print(r.text)

def select_boxno_by_referenceno(referenceno):
    sqlcommand = """create TEMPORARY table if not EXISTS temp1(
                        referenceno VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""
    cur.execute(sqlcommand)
    
    sqlcommand = """insert into temp1 values %s;"""
    cur.execute(sqlcommand % referenceno)
    sqlcommand = """select ib.BOX_NO from temp1 
                    left JOIN order_basic ob on temp1.referenceno = ob.REFERENCE_NO
                    LEFT JOIN inventory_basic ib on ob.ID = ib.ORDER_ID;"""
    
    cur.execute(sqlcommand)                
    for row in cur:
        print(row[0])


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
    referenceno.append("('P41614881611205022')")
    referenceno.append("('P41614881611205022')")
    referenceno.append("('P41614881611205022')")
    
    select_boxno_by_referenceno(','.join(referenceno))
    #关闭mysql
    cur.close()
    conn.close()
