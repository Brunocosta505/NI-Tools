from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path)

for i in wb.sheets():
	if i.name == 'Info': # Skip first sheet
		continue
	print ('Sheet: ' + i.name + ', Rows: ' + repr(i.nrows) + ', Cols: ' + repr(i.ncols))	
	for row in range(i.nrows):
		values = []
		for col in range(i.ncols):
			values.append(i.cell(row,col).value)
		#print (','.join(values))
		#print (repr(values))
	#print ()
	sh = wb.sheet_by_index(1)
	
	for iR in i.nrows:
		for iC in i.ncols:
			cell = sh.cell(iR,iC)
			print (cell)


#
#
#
# print number of sheets
#print ("Number of worksheets: ", wb.nsheets)
#
# print sheet names
#print ("Worksheet name(s):", wb.sheet_names())
#
# get the first worksheet

#
#
#
# read a cell
#cell = sh.cell(0,0)
#print (cell)
#
#
# read a row slice
#print (sh.row_slice(rowx=0, start_colx=0, end_colx=2))