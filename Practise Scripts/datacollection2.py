#tagcloud 试手
import configparser
import pymysql
import codecs

def connect_to_jerry():

    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')

    server = cf.get("JERRY","server")
    port = cf.getint("JERRY","port")
    user = cf.get("JERRY","user")
    password = cf.get("JERRY","password")
    database = cf.get("JERRY","database")
    charset = cf.get("JERRY","charset")
    conn = pymysql.connect(host=server, port=port, user=user, passwd=password, db=database,charset=charset)
    cur = conn.cursor()
    return cur


emmis = connect_to_jerry()
sqlcommand = """SELECT oa.CITY FROM `order_basic` ob left join order_address oa 
on ob.ADDRESS_ID = oa.ID where ob.CREATED_DATE between '2017-04-01' and '2017-04-24';"""

emmis.execute(sqlcommand)

f = codecs.open("city.txt", 'a', 'utf-8')

dict ={}
n = 0
for row in emmis:
    if row[0] in dict:
        dict[row[0]] += 1
    else:
        dict[row[0]] = 1
    n+=1


f.write('total'+'\t'+str(n)+'\r\n')
for key in dict:
    if key is not None:
        f.write(key+'\t'+str(dict[key])+'\r\n')

f.close()
