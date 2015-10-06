from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path) # Open excel
#headers = []
#values = []
#output = []

def GetFreqBand( str ): # Get frequency band
	if str == '': return 0
	band = [int(s) for s in str.split() if s.isdigit()] # Loop through string and get digit
	band = band[0]
	return int(band)

def GetHopMode( str ): # Get hopping mode
	if str == 'Radio Frequency hopping': 
		hop = "RF"
	# Falta BB e No HOP
	else:
		hop = "Other"
	return hop
	
def GetValues( str ): # Get values
	if str is None:
		print ("str is None!!")
		return
	if i.name == 'Site_BTS':
		for c in range(len(str[0])): # Loop through columns
			if Matrix[0][c] == "BCF ID":
				#print (Matrix[0][c])
				col = c
				#print (col)
				
		for r in range(len(str)): # Loop through rows
			if r > 2:
				col2 = Matrix[r][col]
				print (col2)
		return col
	return
	
for i in wb.sheets(): # Loop through sheets
	if i.name == 'Info': # Skip Info sheet
		continue
	
	print ("Sheet: " + i.name + ", Rows: " + repr(i.nrows) + ", Cols: " + repr(i.ncols))
	Matrix = [[0 for x in range(i.ncols)] for x in range(i.nrows)]
	
	for r in range(i.nrows): # Loop through rows
		#values = []
		
		for c in range(i.ncols):  # Loop through columns
			#if r == 0: # Get headers
			cell = i.cell(r, c).value
			#if cell.ctype == XL_CELL_NUMBER:
			#	print ("R: " + repr(int(r)) + ", C: " + repr(int(c)) + ", value:" + str(i.cell(r, c).value))
			#else:
			#	print ("R: " + repr(int(r)) + ", C: " + repr(int(c)) + ", value:" + i.cell(r, c).value)
			Matrix[r][c] = cell
			#elif r > 2: # Get values
				#values.append(i.cell(r, c).value)
		#print (','.join(values))
		#print (repr(values))
	#print (Matrix)
	GetValues(Matrix)
		#if r > 2 and i.name == 'Site_BTS':
		#	hop = GetHopMode(values[21])
		#	band = GetFreqBand(values[33])
		#	# int(x) -> remove decimal places
		#	# repr(x) -> convert int to string
		#	
		#	print ("ZEQC:BCF=" + repr(int(values[10])) + ",BTS=" + repr(int(values[12])) +
		#	",NAME=" + values[3] + ",SEGNAME=" + values[2] + ":CI=" + repr(int(values[16])) +
		#	",BAND=" + repr(band))# +
			#":NCC=" + repr(values[18]))# +
			#",BCC=" + repr(int(values[19])))# +
			#":MCC=" + repr(int(values[14])) +
			#",MNC=" + values[15])# +
			
			#print ("ZEQC:BCF=" + repr(int(values[10])) + ",BTS=" + repr(int(values[12])) +
			#",NAME=" + values[3] + ",SEGNAME=" + values[2] + ":CI=" + repr(int(values[16])) +
			#",BAND=" + int(band) + ":NCC=" + repr(int(values[18])) + ",BCC=" + repr(int(values[19])) +
			#":MCC=" + repr(int(values[14])) +
			#",MNC=" + values[15])# +
			#",LAC=" + repr(int(values[17]))) #+
			#":HOP=" + hop + ",HSN1=" + repr(int(values[23])) + ";")
			#print ('\n')
			#print (output)
			#output = []	
	