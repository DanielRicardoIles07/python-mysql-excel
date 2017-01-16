# -*- coding: utf-8 -*-
import xlwt
import pymysql

#设置基本变量
_host = 'localhost'
_db = '13net'
_user = 'root'
_password = 'root'
_table = 'net_members'
_excel_name = 'tes2t'

#注意编码
conn = pymysql.connect(host=_host,user=_user,password=_password,db=_db,charset='utf8')
cursor = conn.cursor()
count = cursor.execute('select id,username,student_no from %s'%_table);
print('has %s line'%count);

#重置游标
cursor.scroll(0,mode='absolute')

#结果
ret = cursor.fetchall();

#头部
fields = cursor.description
#创建excel
excel = xlwt.Workbook()
#创建工作簿
sheet = excel.add_sheet(_excel_name,cell_overwrite_ok=True)
#写入字段名
for k,v in enumerate(fields):
  sheet.write(0,k,v[0])
#写入数据
for key,value in enumerate(ret):
  for kk,vv in enumerate(value):
    sheet.write(key+1,kk,vv)

excel.save('./%s.xlsx'%_excel_name)
