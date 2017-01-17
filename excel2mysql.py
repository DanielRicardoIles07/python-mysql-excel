# -*- coding: utf-8 -*-
import xlrd
import pymysql

#设置基本变量
_host = 'localhost'
_db = '13net'
_user = 'root'
_password = 'root'
_table = 'net_members_bak'
_excel_name = './tes2t.xlsx'

#open excel
excel = xlrd.open_workbook(_excel_name)
sheet = excel.sheet_by_index(0)

rows = sheet.nrows
cols = sheet.ncols
data = []
fields=''
#创建好要数据,如果第一行是表头的话，从1开始，若第一行就是数据，从0开始
for i in range(1,rows):
  data.append(sheet.row_values(i))

for i in range(0,cols):
    fields = fields+'%s,'
print(fields)
# mysql
conn = pymysql.connect(host=_host,user=_user,password=_password,db=_db,charset='utf8')
cursor = conn.cursor()

#个人觉得最好先创建好表之后来导入数据把，如果要新建的话，也可以在这执行语句新建，但是不建议这么做

#批量插入数据
cursor.executemany("insert into "+_table+" values("+fields[:-1]+");" ,data)
#不要忘记commit
conn.commit()
