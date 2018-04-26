# -*- coding: utf-8 -*-

'''
this script is used for import Mysql`s data to excel
'''
import numpy as np
import pandas as pd
import pymysql

# your Database  config
settings = {"host":"192.168.0.1", "database":"database", "user":"loginuser", "password":"dbpassword", "port":3306, "charset":"utf8"}
# your data source
yourExcelPath = '/tmp/upload/source.csv'
# where to write the data
yourImportPath = '/tmp/upload/improt_data.xlsx'

if yourExcelPath.split('.')[-1] == 'csv':
    source = pd.read_csv(yourExcelPath, encoding='gbk', dtype={'SKU':str}).dropna(axis=0)
elif yourExcelPath.split('.')[-1] == 'xlsx':
    source = pd.read_excel(yourExcelPath, encoding='gbk', dtype={'SKU':str}).dropna(axis=0)
else:
    print('can`t read the source file')

where = '","'.join(source['SKU'])
dbobj= pymysql.connect(host=settings['host'], database=settings['database'], user=settings['user'],
					password=settings['password'], port=settings['port'], charset=settings['charset'])
selectsql = 'select goods_sn,product_character_ids from p_custom_declaration where goods_sn in ("%s")' % where
data = pd.read_sql(selectsql, dbobj)
excel = pd.ExcelWriter(yourImportPath)
data.to_excel(excel)
excel.save()
print('finish')