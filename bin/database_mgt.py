import pymssql

def get_datbase_connect():
    connect = pymssql.connect('192.168.2.17', 'sa', 'Dwwy001', 'SoochowDb')  # 建立连接
    if connect:
        print("连接成功!")
    return connect

def close_connect(connect):
    connect.close()

def get_buildings(connect):
    cursor = connect.cursor()
    sql = "select * from Buildings"

    cursor.execute(sql)
    row = cursor.fetchone()
    while row:
        print(row)
        row = cursor.fetchone()

    cursor.close()


connect = get_datbase_connect()
get_buildings(connect)
close_connect(connect)