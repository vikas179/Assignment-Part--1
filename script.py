import xlrd
import sys
import xlwt
from xlutils.copy import copy
ppdbook = xlrd.open_workbook('PPD.xlsx')
codbook = xlrd.open_workbook('COD.xlsx')
print(sys.argv[1])
workbook = xlrd.open_workbook(sys.argv[1])

worksheet_invoice = workbook.sheet_by_index(0)

rows = worksheet_invoice.nrows
cols = worksheet_invoice.ncols

# Cols 0-10 invoice 1
# Cols 12-22 invoice 2

invoice_1 = dict()
invoice_2 = dict()

i = 4
for i in range(4,rows):
	if i == 4:
		#print(worksheet_invoice.cell_value(i, 1))
		invoice_1['SELLER_GSTIN'] = worksheet_invoice.cell_value(i, 1).split('- ')[1]
		invoice_2['SELLER_GSTIN'] = worksheet_invoice.cell_value(i, 13).split('- ')[1]
		

	elif i == 7:
		invoice_1['INVOICE_NUMBER'] = worksheet_invoice.cell_value(i, 1)
		invoice_1['PRODUCT'] = worksheet_invoice.cell_value(i, 3).split()[0]
		invoice_1['INVOICE_DATE'] = worksheet_invoice.cell_value(i, 8)

		invoice_2['INVOICE_NUMBER'] = worksheet_invoice.cell_value(i, 13)
		invoice_2['PRODUCT'] = worksheet_invoice.cell_value(i, 15).split()[0]
		invoice_2['INVOICE_DATE'] = worksheet_invoice.cell_value(i, 20)

	elif i == 10:
		invoice_1['ORDER_NUMBER'] = worksheet_invoice.cell_value(i, 1)
		invoice_2['ORDER_NUMBER'] = worksheet_invoice.cell_value(i, 13)

	elif i == 13:
		invoice_1['CONSIGNEE'] = worksheet_invoice.cell_value(i, 2)
		invoice_2['CONSIGNEE'] = worksheet_invoice.cell_value(i, 14)

	elif i == 14:
		invoice_1['CONSIGNEE_ADDRESS1'] = worksheet_invoice.cell_value(i, 2)
		invoice_2['CONSIGNEE_ADDRESS1'] = worksheet_invoice.cell_value(i, 14)
		invoice_1['DESTINATION_CITY'] = invoice_1['CONSIGNEE_ADDRESS1'].split(',')[-1]
		invoice_2['DESTINATION_CITY'] = invoice_2['CONSIGNEE_ADDRESS1'].split(',')[-1]

	elif i == 15:
		invoice_1['MOBILE'] = worksheet_invoice.cell_value(i, 2).split(': ')[1]
		invoice_2['MOBILE'] = worksheet_invoice.cell_value(i, 14).split(': ')[1]

	elif i == 16:
		invoice_1['STATE'] = worksheet_invoice.cell_value(i, 3)
		invoice_2['STATE'] = worksheet_invoice.cell_value(i, 15)

	elif i == 17:
		invoice_1['STATE_CODE'] = int(worksheet_invoice.cell_value(i, 3))
		invoice_2['STATE_CODE'] = int(worksheet_invoice.cell_value(i, 15))
		

	elif i == 22:
		invoice_1['ITEM_DESCRIPTION'] = worksheet_invoice.cell_value(i, 1)
		invoice_1['PIECES'] = int(worksheet_invoice.cell_value(i, 6))
		invoice_1['GST_TAX_BASE'] = worksheet_invoice.cell_value(i, 7)

		invoice_2['ITEM_DESCRIPTION'] = worksheet_invoice.cell_value(i, 13)
		invoice_2['PIECES'] = int(worksheet_invoice.cell_value(i, 18))
		invoice_2['GST_TAX_BASE'] = worksheet_invoice.cell_value(i, 19)

	elif i == 27:
		invoice_1['GST_TAX_TOTAL'] = worksheet_invoice.cell_value(i, 9)
		invoice_1['GST_TAX_NAME'] = 'HR '+ worksheet_invoice.cell_value(i, 8)

		invoice_2['GST_TAX_TOTAL'] = worksheet_invoice.cell_value(i, 21)
		invoice_2['GST_TAX_NAME'] = 'HR '+ worksheet_invoice.cell_value(i, 20)

	elif i == 28:
		invoice_1['DECLARED_VALUE'] = worksheet_invoice.cell_value(i, 9)
		invoice_2['DECLARED_VALUE'] = worksheet_invoice.cell_value(i, 21)



'''
	elif i == 15:
		invoice
		'''
if 'PPD' in invoice_1['PRODUCT']:
	invoice_1['COLLECTABLE_VALUE'] = 0
elif 'COD' in invoice_2['PRODUCT']:
	invoice_2['COLLECTABLE_VALUE'] = invoice_2['DECLARED_VALUE']

if 'PPD' in invoice_2['PRODUCT']:
	invoice_2['COLLECTABLE_VALUE'] = 0
elif 'COD' in invoice_2['PRODUCT']:
	invoice_2['COLLECTABLE_VALUE'] = invoice_2['DECLARED_VALUE']


ppdwrite = copy(xlrd.open_workbook('PPD.xlsx'))
codwrite = copy(xlrd.open_workbook('COD.xlsx'))

if 'PPD' in invoice_1['PRODUCT']:
	ppdsheet_read = ppdbook.sheet_by_index(0)
	for i in range(ppdsheet_read.nrows):
		if ppdsheet_read.cell_value(i, 1) == 'Unused':
			invoice_1['AWB_NUMBER'] = ppdsheet_read.cell_value(i, 0)
			ppdsheet_write = ppdwrite.get_sheet(0)
			ppdsheet_write.write(i, 1, 'Used')
			ppdwrite.save('PPD.xlsx')
			break

elif 'COD' in invoice_1['PRODUCT']:
	codsheet_read = codbook.sheet_by_index(0)
	for i in range(codsheet_read.nrows):
		if codsheet_read.cell_value(i, 1) == 'Unused':
			invoice_1['AWB_NUMBER'] = codsheet_read.cell_value(i, 0)
			codsheet_write = codwrite.get_sheet(0)
			codsheet_write.write(i, 1, 'Used')
			codwrite.save('COD.xlsx')
			break

if 'PPD' in invoice_2['PRODUCT']:
	ppdsheet_read = ppdbook.sheet_by_index(0)
	for i in range(ppdsheet_read.nrows):
		if ppdsheet_read.cell_value(i, 1) == 'Unused':
			invoice_2['AWB_NUMBER'] = ppdsheet_read.cell_value(i, 0)
			ppdsheet_write = ppdwrite.get_sheet(0)
			ppdsheet_write.write(i, 1, 'Used')
			ppdwrite.save('PPD.xlsx')
			break

elif 'COD' in invoice_2['PRODUCT']:
	codsheet_read = codbook.sheet_by_index(0)
	for i in range(codsheet_read.nrows):
		if codsheet_read.cell_value(i, 1) == 'Unused':
			invoice_2['AWB_NUMBER'] = codsheet_read.cell_value(i, 0)
			codsheet_write = codwrite.get_sheet(0)
			codsheet_write.write(i, 1, 'Used')
			codwrite.save('COD.xlsx')
			break



fields = ['AWB_NUMBER','ORDER_NUMBER','PRODUCT','CONSIGNEE','CONSIGNEE_ADDRESS1','DESTINATION_CITY','STATE','MOBILE', 
'ITEM_DESCRIPTION','PIECES','COLLECTABLE_VALUE','DECLARED_VALUE','INVOICE_NUMBER' , 'INVOICE_DATE','SELLER_GSTIN', 'GST_TAX_NAME',
'GST_TAX_BASE', 'GST_TAX_TOTAL']


output = copy(xlrd.open_workbook('report.xlsx'))
output_read = xlrd.open_workbook('report.xlsx')
outputsheet_write = output.get_sheet(0)
rows = output_read.sheet_by_index(0).nrows
row = outputsheet_write.row(rows)
for i in range(len(fields)):
	row.write(i,invoice_1[fields[i]])

rows+=1
row = outputsheet_write.row(rows)
for i in range(len(fields)):
	row.write(i,invoice_2[fields[i]])


output.save('report.xlsx')

print(invoice_1)
print('==============')
print(invoice_2)
