from xlrd import open_workbook

path = 'test.xlsx'
wb = open_workbook(path) # Open excel
numHead = 0 # Number of headers
KW = ['BCF ID', 'BTS ID', 'BTS NAME', 'SEG NAME', 'CI', 'BAND', 'NCC', 'BCC', 'MCC', 'MNC', 'LAC', 'Hopping Mode', 'HSN 1', 'HSN 2'] # Parameters key words
rowKW = 1 # Line to search key works
tmp = -1

class Sheet:
    def __init__(self, r, c): # Initialization
        self.data = [[0 for x in range(c)] for x in range(r)] # Data handler
        self.keyWords = [[tmp for x in range(len(KW))] for x in range(2)] # Parameters and index

        for i in range(len(self.keyWords)): # Fill array with parameters
            for j in range(len(self.keyWords[i])):
                if i == 0:
                    self.keyWords[i][j] = KW[j] 

    def readValue(self, r, c, value): # Read and save value to data array
        self.data[r][c] = value
        return self.getNumHead(value)
            
    def getNumHead(self, value): # Check number of headers
        if numHead != 0:
            return numHead
        else:
            nHead = 0
            if c == 0:
                if self.hasNumbers(value):
                    nHead = r
            return nHead
    
    def hasNumbers(self, inputString): # Check if contains digit
        return any(char.isdigit() for char in inputString)

    def kwSearch(self, sheet): # Get parameters indexes
        if sheet == 'Site_BTS':
            for j in range(len(self.data[0])): # Loop through data columns
                #print (self.data[j])
                for i in range(len(self.keyWords[0])): # Loop through parameters array
                    if self.data[rowKW - 1][j] == self.keyWords[0][i]:
                        self.keyWords[1][i] = j
            #print (self.keyWords)

    def printMML(self, sheet):
        if sheet == 'Site_BTS':
            for j in range(numHead, len(self.data)): # Loop through data rows
                #print (self.keyWords)
                #print (self.data[j][self.keyWords[1][5]])
                band = self.GetFreqBand(str(self.data[j][self.keyWords[1][5]]))
                hop = self.GetHopMode(str(self.data[j][self.keyWords[1][11]]))
                
                print ("ZEQC:BCF=" + str(int(self.data[j][self.keyWords[1][0]]))
                + ",BTS=" + str(int(self.data[j][self.keyWords[1][1]]))
                + ",NAME=" + str(self.data[j][self.keyWords[1][2]])
                + ",SEGNAME=" + str(self.data[j][self.keyWords[1][3]]) # // REVIEW IF NEEDED
                + ":CI=" + str(int(self.data[j][self.keyWords[1][4]]))
                + ",BAND=" + str(band)
                + ":NCC=" + str(int(self.data[j][self.keyWords[1][6]]))
                + ",BCC=" + str(int(self.data[j][self.keyWords[1][7]]))
                + ":MCC=" + str(int(self.data[j][self.keyWords[1][8]]))
                + ",MNC=" + str(int(self.data[j][self.keyWords[1][9]]))
                + ",LAC=" + str(int(self.data[j][self.keyWords[1][10]]))
                + ":HOP=" + str(hop)
                + ",HSN1=" + str(int(self.data[j][self.keyWords[1][12]]))
                + ",HSN2=" + str(int(self.data[j][self.keyWords[1][13]]))
                + ";")
                
    def GetFreqBand(self, value): # Get frequency band
        if value == "": return "[!!!]"
        band1 = [int(s) for s in value.split() if s.isdigit()] # Loop through string and get digit
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
    print ("Sheet: " + i.name + ", Rows: " + repr(i.nrows) + ", Cols: " + repr(i.ncols))   
    data = Sheet(i.nrows, i.ncols)
    
    for r in range(i.nrows): # Loop through rows
        for c in range(i.ncols):  # Loop through columns
            numHead = data.readValue(r, c, i.cell(r, c).value)
    
        if data.keyWords[1][0] == tmp: # Get parameters indexes
            data.kwSearch(i.name)
            
    data.printMML(i.name)
    
    