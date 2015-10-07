from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path) # Open excel
numHead = 3 # Number of headers // REVIEW: might be automated through function
KW = ['BCF ID', 'BTS ID', 'BTS NAME', 'SEG NAME', 'CI', 'BAND', 'NCC', 'BCC', 'MCC', 'MNC', 'LAC', 'Hopping Mode', 'HSN 1', 'HSN 2'] # Key words
rowKW = 1 # Line to search key works

tmp = -6

class Sheet:
    def __init__(self, r, c):
        self.data = [[0 for x in range(c)] for x in range(r)] # Data handler
        self.keyWords = [[tmp for x in range(len(KW))] for x in range(2)] # Key words and index
        
        for i in range(len(self.keyWords)):
            for j in range(len(self.keyWords[i])):
                if i == 0:
                    self.keyWords[i][j] = KW[j] # Fill array with key words

    def readValue(self, r, c, value): # Read and save value to data array
        self.data[r][c] = value
    
    def kwSearch(self, sheet): # Get key words index
        if sheet == 'Site_BTS':
            for j in range(len(self.data[0])): # Loop through data columns
                for i in range(len(self.keyWords[0])): # Loop through key words array
                    if self.data[rowKW - 1][j] == self.keyWords[0][i]:
                        self.keyWords[1][i] = j
            #print (self.keyWords)

    def printMML(self, sheet):
        if sheet == 'Site_BTS':
            for j in range(numHead, len(self.data)): # Loop through data rows
                #print (self.keyWords)
                band = self.GetFreqBand(str(self.data[j][self.keyWords[1][5]]))
                hop = self.GetHopMode(str(self.data[j][self.keyWords[1][11]]))
                
                print ("ZEQC:BCF=" + str(int(self.data[j][self.keyWords[1][0]]))
                + ",BTS=" + str(int(self.data[j][self.keyWords[1][1]]))
                + ",NAME=" + str(self.data[j][self.keyWords[1][2]])
                + ",SEGNAME=" + str(self.data[j][self.keyWords[1][3]])
                + ":CI=" + str(int(self.data[j][self.keyWords[1][4]]))
                + ",BAND=" + str(band)
                + ":NCC=" + str(int(self.data[j][self.keyWords[1][6]]))
                + ",BCC=" + str(int(self.data[j][self.keyWords[1][7]]))
                + ":MCC=" + str(int(self.data[j][self.keyWords[1][8]]))
                + ",MNC=" + str(int(self.data[j][self.keyWords[1][9]]))
                + ",LAC=" + str(int(self.data[j][self.keyWords[1][10]]))
                + ":HOP=" + str(hop)
                + ",HSN1=" + str(int(self.data[j][self.keyWords[1][12]]))
                + ",HSN2=" + str(int(self.data[j][self.keyWords[1][13]])) # // REVIEW
                + ";")
    
    def GetFreqBand(self, value): # Get frequency band
        if value == "": return "[!!!]"
        band1 = [int(s) for s in value.split() if s.isdigit()] # Loop through string and get digit only
        band = band1[0]
        return band
    
    def GetHopMode(self, value): # Get hopping mode
        if value == "Radio Frequency hopping": 
            hop = "RF"
        # // REVIEW: missing if for bb and no hopping
        else:
            hop = "[!!!]"
        return hop
    
for i in wb.sheets(): # Loop through sheets
    if i.name == "Info": # Skip Info sheet
        continue
    
    print ("Sheet: " + i.name + ", Rows: " + repr(i.nrows) + ", Cols: " + repr(i.ncols))   
    data = Sheet(i.nrows, i.ncols)
    
    for r in range(i.nrows): # Loop through rows
        for c in range(i.ncols):  # Loop through columns
            data.readValue(r, c, i.cell(r, c).value)
    
        if data.keyWords[1][0] == tmp: # Get values once
            data.kwSearch(i.name)
            
    data.printMML(i.name)
    
    