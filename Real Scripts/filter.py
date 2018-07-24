import pandas as pd
import configparser
import pypyodbc

SQL_CONFIG_FILE = "sql.ini"


def connect_to_SQLServer(str):
    ''' connect to MYSQL database

    Args:
        str: EMMIS,OMS,TMS
    Return:
        Connection, can be used as conn.cursor() and conn.cursor("select count(1) from ...")
    '''

    cf = configparser.ConfigParser()
    cf.read(SQL_CONFIG_FILE)
    driver = cf.get(str, "driver")
    server = cf.get(str, "server")
    user = cf.get(str, "user")
    password = cf.get(str, "password")
    database = cf.get(str, "database")
    connection_string = "Driver={" + driver + "};"
    connection_string += "Server=" + server + ";"
    connection_string += "UID=" + user + ";"
    connection_string += "PWD=" + password + ";"
    connection_string += "Database=" + database + ";"
    cnxn = pypyodbc.connect(connection_string)
    return cnxn


def emmis_get_box_no(box_no):
    '''Get all finished box_no

    Arguments:
        box_no {list} -- all box no that has invoice
    Return:
        res {set} -- all box no that is finished in emmis
    '''
    print("start query from emmis")
    res = set()
    composit_list = [box_no[x:x + 1000] for x in range(0, len(box_no), 1000)]

    with connect_to_SQLServer("EMMIS").cursor() as emmis_cursor:
        for thousand_box_no in composit_list:
            sqlcommand = '''select cnum from client_rec where nstate = 3 and cnum in ('{}');'''.format(
                "','".join(thousand_box_no))

            emmis_cursor.execute(sqlcommand)

            if emmis_cursor:
                for line in emmis_cursor:
                    res.add(line[0])

    print("query end")
    return res


if __name__ == "__main__":
    df = pd.read_csv("june.csv")
    box_no = list(df["箱号"])
    finished_box_no = emmis_get_box_no(box_no)
    finished_data = df[df["箱号"].isin(finished_box_no)]
    finished_data.to_csv("finished.csv")
