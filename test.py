from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path)
headers = []
values = []
#output = []

for i in wb.sheets():
	if i.name == 'Info': # Skip Info sheet
		continue
	print ("Sheet: " + i.name + ", Rows: " + repr(i.nrows) + ", Cols: " + repr(i.ncols))	
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
				
			if values[33] != '': # Get frequency band
				band = [int(s) for s in values[33].split() if s.isdigit()]	
				
			print ("ZEQC:BCF=" + repr(int(values[10])) + ",BTS=" + repr(int(values[12])) +
			",NAME=" + values[3] + ",SEGNAME=" + values[2] + ":CI=" + repr(int(values[16])) +
			",BAND=" + int(band) + ":NCC=" + repr(int(values[18])) + ",BCC=" + repr(int(values[19])) +
			":MCC=" + repr(int(values[14])) + ",MNC=" + repr(int(values[15])) + ",LAC=" + repr(int(values[17])) +
			":HOP=" + str(hop) + ",HSN1=" + repr(int(values[23])) + ";")
			#print ('\n')
			#print (output)
			#output = []	
	