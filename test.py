from xlrd import open_workbook

path = 'test.xlsx' #Excel file path
KW = ['BCF ID', 'BTS ID', 'BTS NAME', 'SEG NAME', 'CI', 'BAND', 'NCC', 'BCC', 'MCC', 'MNC', 'LAC', 'Hopping Mode', 'HSN 1', 'HSN 2'] # Parameter key words
rowKW = 1 # Line to search key works
iniValue = -1 #Initial value of array

class Sheet:
    def __init__(self, r, c): # Initialization
        self.data = [[0 for x in range(c)] for x in range(r)] # Data handler
        self.keyWords = [[iniValue for x in range(len(KW))] for x in range(2)] # Parameters and index
        self.numHead = 0 # Number of headers

        for i in range(len(self.keyWords)): # Fill array with parameters
            for j in range(len(self.keyWords[i])):
                if i == 0:
                    self.keyWords[i][j] = KW[j] 

    def readValue(self, r, c, value): # Read and save value to data array
        self.data[r][c] = value
        return self.getNumHead(value)
            
    def getNumHead(self, value): # Check number of headers
        if self.numHead != 0:
            return self.numHead
        else:
            if c == 0:
                if self.hasNumbers(value):
                    self.numHead = r
    
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
            for j in range(self.numHead, len(self.data)): # Loop through data rows
                #print (self.keyWords)
                #print (self.data[j][self.keyWords[1][5]])
                
                           
                output = "ZEQC:BCF=" + str(int(self.data[j][self.keyWords[1][0]]))
                output += ",BTS=" + str(int(self.data[j][self.keyWords[1][1]]))
                output += ",NAME=" + str(self.data[j][self.keyWords[1][2]])
                output += ",SEGNAME=" + str(self.data[j][self.keyWords[1][3]]) # // REVIEW IF NEEDED
                
                output += ":CI=" + str(int(self.data[j][self.keyWords[1][4]]))
                band = self.GetFreqBand(str(self.data[j][self.keyWords[1][5]]))
                output += ",BAND=" + str(band)
                
                output += ":NCC=" + str(int(self.data[j][self.keyWords[1][6]]))
                output += ",BCC=" + str(int(self.data[j][self.keyWords[1][7]]))
                
                output += ":MCC=" + str(int(self.data[j][self.keyWords[1][8]]))
                output += ",MNC=" + str(int(self.data[j][self.keyWords[1][9]]))               
                output += ",LAC=" + str(int(self.data[j][self.keyWords[1][10]]))
                
                hop = self.GetHopMode(str(self.data[j][self.keyWords[1][11]]))
                output += ":HOP=" + str(hop)
                
                hsn1 = self.keyWords[1][12]
                if hsn1 != iniValue:
                    output += ",HSN1=" + str(int(self.data[j][hsn1]))
                
                hsn2 = self.keyWords[1][13]
                if hsn2 != iniValue:
                    output += ",HSN2=" + str(int(self.data[j][hsn2]))
                
                output += ";"
                print(output)
                
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

        
wb = open_workbook(path) # Open excel        
for i in wb.sheets(): # Loop through sheets
    print ("Sheet: " + i.name + ", Rows: " + repr(i.nrows) + ", Cols: " + repr(i.ncols))   
    data = Sheet(i.nrows, i.ncols)
    
    for r in range(i.nrows): # Loop through rows
        for c in range(i.ncols):  # Loop through columns
            data.readValue(r, c, i.cell(r, c).value)
    
        if data.keyWords[1][0] == iniValue: # Get parameters indexes
            data.kwSearch(i.name)
            
    data.printMML(i.name)
    
    