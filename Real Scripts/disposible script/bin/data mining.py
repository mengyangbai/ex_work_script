

import configparser
import pypyodbc

SQL_CONFIG_FILE=r'..\config\sql.ini'
def connect_to_SQLServer(str):
    ''' connect to MYSQL database
    
    Args:
        str: EMMIS,OMS,TMS
    Return:
        Connection, can be used as conn.cursor() and conn.cursor("select count(1) from ...")
    '''

    cf = configparser.ConfigParser()
    cf.read(SQL_CONFIG_FILE)
    driver = cf.get(str,"driver")
    server = cf.get(str,"server")
    user = cf.get(str,"user")
    password = cf.get(str,"password")
    database = cf.get(str,"database")
    connection_string = "Driver={"+driver+"};"
    connection_string += "Server="+server+";"
    connection_string += "UID="+user+";"
    connection_string += "PWD="+password+";"
    connection_string += "Database="+database+";"
    cnxn = pypyodbc.connect(connection_string)
    return cnxn
    
if __name__=='__main__':
    with open('1.txt') as f, open('2.txt','a') as t,connect_to_SQLServer("EMMIS").cursor() as ecur:
        n=0
        a = set()
        for line in f:
            a.add(line.rstrip('\r\n'))
            n+=1
            if n == 1000:
                string = "','".join(a)
                sqlcommand = '''select cnum from client_rec where cnum in ('{}') and nstate = 3;'''.format(string)
                ecur.execute(sqlcommand)
                for row in ecur:
                    t.write(row[0]+'\n')
                n=0
                a.clear()
                
        string = "','".join(a)
        sqlcommand = '''select cnum from client_rec where cnum in ('{}') and nstate = 3;'''.format(string)
        ecur.execute(sqlcommand)
        for row in ecur:
            t.write(row[0]+'\n')
        n=0
        a.clear()
        
        
                
                
                
                
                