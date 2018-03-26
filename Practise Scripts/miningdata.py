#datamining 试手
#问题很多，暂时弃用，到时用numpy和pandas
import io,sys
import configparser
import pypyodbc
import pymysql

def connect_to_emmis():

    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')
    #先连接sqlserver
    print("连接到SQLSERVER")
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
    newCnxn = pypyodbc.connect(connection_string)
    return newCnxn
    
def connect_to_JERRY():

    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')
    #先连接Jerry
    print("连接到JERRY")
    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    return conn

#解决输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')    
emmis = connect_to_emmis()
emmis_cursor = emmis.cursor()
jerry = connect_to_JERRY()
jerry_cursor = jerry.cursor()
sqlcommand="""select cr.cnum from client_rec cr where cr.nstate = 3 
	and cr.dsysdate between '2017-03-01' and '2017-04-01' order by cr.cnum;"""
emmis_cursor.execute(sqlcommand)

sqlcommand = """create TEMPORARY table if not EXISTS temp1(
                        BOX_NO VARCHAR(250)
                    )DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""
jerry_cursor.execute(sqlcommand)

list=[]
values=""
sqlcommand="insert into temp1 values "
for row in emmis_cursor:
    if len(list) == 100:
        values = values[1:]
        sqlcommand = sqlcommand + values;
        sqlcommand +=";"
        jerry_cursor.execute(sqlcommand)
        list=[]
        sqlcommand = "insert into temp1 values "
        values = ",('"+row[0]+"')"
    else:
        list.append(row[0])
        values +=",('"+row[0]+"')"

values = values[1:]
sqlcommand = sqlcommand + values;
sqlcommand +=";"        
jerry_cursor.execute(sqlcommand)

        

sqlcommand="""select oa.`NAME`,oa.MOBILE,ib.BOX_NO 
                from order_address oa LEFT JOIN 
                order_basic ob on oa.id = ob.ADDRESS_ID
                left join inventory_basic ib on ib.ORDER_ID = ob.ID
                left join temp1 on temp1.BOX_NO = ib.BOX_NO;"""    

jerry_cursor.execute(sqlcommand)

dict={}

for row in jerry_cursor:
    print(row[0]+","+row[1]+","+row[2])


        

jerry_cursor.close()
jerry.close()
emmis_cursor.close()
emmis.close()