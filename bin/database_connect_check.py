import pymssql

connect = pymssql.connect('192.168.2.17', 'sa', 'Dwwy001', 'SoochowDb')  # 建立连接
if connect:
    print("连接成功!")

cursor = connect.cursor()
sql = "select * from Buildings"

cursor.execute(sql)
row = cursor.fetchone()
while row:
    print(row)
    row = cursor.fetchone()

cursor.close()
connect.close()