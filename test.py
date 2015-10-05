from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path)
headers = []
values = []
output = []

for i in wb.sheets():
	if i.name == 'Info': # Skip Info sheet
		continue
	print ('Sheet: ' + i.name + ', Rows: ' + repr(i.nrows) + ', Cols: ' + repr(i.ncols))	
	for r in range(i.nrows):
		values = []
		for c in range(i.ncols):
			if r == 0: # Get headers
				headers.append(i.cell(r, c).value)
			elif r > 2: # Get values
				values.append(i.cell(r, c).value)
		#print (','.join(values))
		#print (repr(values))
		#print (repr(values))
		if r > 2 and i.name == 'Site_BTS':
			if values[21] == 'Radio Frequency hopping':
				hop = "RF"
			else:
				hop = "Other"
			output.append('ZEQC:BCF=' + repr(values[10]) + ',BTS=' + repr(values[12]) +
			',NAME=' + values[3] + ',SEGNAME=' + values[2] + ':CI=' + repr(values[16]) + 
			',BAND=' + values[33] + ':NCC=' + repr(values[18]) + ',BCC=' + repr(values[19]) + 
			':MCC=' + repr(values[14]) + ',MNC=' + repr(values[15]) + ',LAC=' + repr(values[17])) #+
			#output.append('\n')
			#':HOP=' + hop + ',HSN1=' + values[23] + ';')
			#print ('\n')
			print (output)
			output = []	
	

	
#sh = wb.sheet_by_index(1)	
#for (i, values) in enumerate(values):
#    print (i, values)
			#cell = sh.cell(iR,iC)
			#print (cell)

# print number of sheets
#print ("Number of worksheets: ", wb.nsheets)
#
# print sheet names
#print ("Worksheet name(s):", wb.sheet_names())
#
# get the first worksheet


# read a cel
#cell = sh.cell(0,0)
#print (cell)
#
#
# read a row slice
#print (sh.row_slice(rowx=0, start_colx=0, end_colx=2))