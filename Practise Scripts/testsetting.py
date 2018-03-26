#测试配置文件
import os
import configparser
import pypyodbc
import pymysql

cf = configparser.ConfigParser()
cf.read('config\\sql.ini')
print(cf.sections())
driver = cf.get("EMMIS","driver")
server = cf.get("EMMIS","server")
user = cf.get("EMMIS","user")
password = cf.get("EMMIS","password")
database = cf.get("EMMIS","database")

connection_string = "Driver={"+driver+"};"
connection_string += "Server="+server+";"
connection_string += "UID="+user+";"
connection_string += "PWD="+password+";"
connection_string += "Database="+database+";"

cnxn = pypyodbc.connect(connection_string)

cursor = cnxn.cursor()
print("haha")

server = cf.get("UIS","server")
port = cf.getint("UIS","port")
user = cf.get("UIS","user")
password = cf.get("UIS","password")
database = cf.get("UIS","database")
charset = cf.get("UIS","charset")
conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
cur = conn.cursor()



#关闭mysql
cur.close()
conn.close()
#关闭sql server
cursor.close()
cnxn.close()
    