fields = ['AWB_NUMBER','ORDER_NUMBER','PRODUCT','CONSIGNEE','CONSIGNEE_ADDRESS1','DESTINATION_CITY','STATE','MOBILE', 
'ITEM_DESCRIPTION','PIECES','COLLECTABLE_VALUE','DECLARED_VALUE','INVOICE_NUMBER' , 'INVOICE_DATE','SELLER_GSTIN', 'GST_TAX_NAME',
'GST_TAX_BASE', 'GST_TAX_TOTAL']

import xlwt
import os
from glob import glob

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Sheet_1')
workbook.save('report.xlsx')
row = sheet.row(0)
for i in range(len(fields)):
	workbook.get_sheet(0).write(0,i, fields[i])

workbook.save('report.xlsx')

files = glob('invoice*.xlsx')
print(files)
#print(type(files[0]))
for file in files:
	os.system('python script.py '+file)

