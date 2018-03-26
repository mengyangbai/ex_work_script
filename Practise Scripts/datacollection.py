#tagcloud 试手
import configparser
import pypyodbc
import codecs

def connect_to_emmis():

    cf = configparser.ConfigParser()
    cf.read('config\\sql.ini')
    #先连接sqlserver
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
    return newCnxn.cursor()


emmis = connect_to_emmis()
sqlcommand = """SELECT cd.cinfo FROM [dbo].[client_rec] cr
                    LEFT JOIN check_detail cd on cr.irid = cd.irid 
                    where cr.nstate = 3 and cr.dsysdate BETWEEN '2017-01-01' and '2017-04-01'
                    and cd.npos>= 100;"""

emmis.execute(sqlcommand)

f = codecs.open("tagcloud.txt", 'a', 'utf-8')

for row in emmis:
    f.write(row[0]+'\r\n')

f.close()
