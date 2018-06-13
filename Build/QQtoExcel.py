import sys
import pandas as pd
import xlwings as xw
import re

def getConv(file):
    #Import conversion table from within excel MQ, set index to PANEL, drop extraneous col
    sh = xw.books(file).sheets['CONVERSIONS']
    conv=pd.DataFrame(sh.range('A1:Y10').value)
    conv.columns = conv.iloc[0]
    conv.set_index('PANEL', inplace=True)
    conv.drop(['PANEL'], inplace=True)
    print(conv)
    return conv

def prepInput(df, num, pan):    
    #Initialize conversion df: add SKU column, and fill out Typ column with panel type
    itemsConv = df
    itemsConv['Type'] = pan
    itemsConv['SKU'] = "" 

    #Loop to turn all Panel Names in the Item column to 'PAN' Also sums up LF of panels and lumps into one column
    numPanCols = 0 
    linFt = 0
    for i in range(num):
    
        if itemsConv.loc[i, 'Type'] == itemsConv.loc[i, 'Item']:
            itemsConv.loc[i, 'Item'] = 'PAN'
            linFt = linFt + itemsConv.loc[i, 'Qty']
            if numPanCols != 0:
                itemsConv.drop(itemsConv.index[i], axis=0, inplace=True)
            numPanCols += 1
       
    #Update Qty of panels to total linFt iterated above
    itemsConv.loc[0,'Qty'] = linFt
    return itemsConv

def QQtoExcel(name):
    print("QQtoxL start") #debug line
    
    mqVersion = 'MultiQuoter2018v0' #This is the filename for current MQ
    
    #Format and split input from csv into item list and pan list
    output = QQtoPD(name)
    itemsRaw = output[0]
    pansRaw = output[1]

    #Prep conv and itemsConv
    numItems = len(itemsRaw['Item'])
    panType = itemsRaw['Item'].values[0]
    conv = getConv(mqVersion)
    itemsConv = prepInput(itemsRaw, numItems, panType)
    linFt = itemsConv.loc[0,'Qty']

    
    #Add screws (round up to 250)
    numScrews = linFt * conv.loc[panType.upper(), 'NUMSCREW']
    bags = round(numScrews/250) + 1
    numScrews = bags * 250
    itemsConv.loc[numItems + 1] = [numScrews, 'SCREW', panType,'']

    
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Loop to convert SKUs
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type'].upper()
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass

    
    #--Add other items, Z flashing, screws, butyl, etc--
    # Check if standing seam (sSeam = 1)
    if itemsConv.loc[0,'SKU'] == 'GL' or itemsConv.loc[0,'SKU'] =='GS' or itemsConv.loc[0,'SKU'] =='DL':
            sSeam = 1
    else:
        sSeam = 0
    
    # For sSeam, add Z flashing to match # of HC, RC, EF, SW, and EW
    # For sSeam, add PS to match # of PV, TF
    # For sSeam, add pancakes for 
    
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    
    #--Loop to match SKUs...this may not need an actual loop
    
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Loop to convert SKUs
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type'].upper()
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass
    
    #Loop to count ZF needs and PS needs and Butyl
    numZF = 0
    numPS = 0
    numButyl = 0
    
    if sSeam == 1:
        for i in range(numItems):
            temp = itemsConv.loc[i,'Item']
            if temp == 'HC' or temp =='RC' or temp== 'EF' or temp == 'SW' or temp == 'EW':
                numZF = numZF + itemsConv.loc[i, 'Qty']
            if temp == 'PV' or temp == 'TF':
                numPS = numPS + itemsConv.loc[i, 'Qty']
        itemsConv.loc[numItems + 1] = [numZF,'ZF', panType, '']
        itemsConv.loc[numItems + 2] = [numPS,'PS', panType, '']
        itemsConv.loc[numItems + 3] = [round((numPS + numZF)/5)+1,'BUTYL', panType, 'BUTYL']

    if sSeam == 0:
        for i in range(numItems):
            temp = itemsConv.loc[i,'Item']
            if temp== 'EF' or temp == 'SW' or temp == 'PV' or temp == 'TF':
                numButyl = numButyl+ itemsConv.loc[i, 'Qty']
        itemsConv.loc[numItems + 1] = [round((numButyl)/5)+1,'BUTYL', panType, 'BUTYL']
        
    #Loop to convert SKUs
    #Reset index and get numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type'].upper()
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass
        
    print(itemsConv)    

    # --Write output from ItemsConv to Excel--
    # This is the xlwings version
    mqVersion = 'MultiQuoter2018v0' #This is the filename for current MQ
    # --Write output from ItemsConv to Excel--

    #designate sheet to write to
    sh = xw.books(mqVersion).sheets['MULTIQUOTER']
    
    #Reset index and get numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Write data from itemsConv to the designated sheet
    for i in range(numItems):
        tempQty = itemsConv['Qty'].values[i]
        tempCode = itemsConv['SKU'].values[i]
        if i == 0:
            sh.range('B14').value = tempQty
            sh.range('C14').value = tempCode
        if i > 0:
            tempCol = 16 + i
            tempQtyLoc = 'B' + str(tempCol)
            tempCodeLoc = 'C' + str(tempCol)
            sh.range(tempQtyLoc).value = tempQty
            sh.range(tempCodeLoc).value=tempCode
    #extract QQ number
    QQID = droppedFile[:-4] #drop ".csv"
    m = re.search(r"[\d]*[-][A-z]{2}[-][\w]*", QQID)
    if m:
        QQID = m.group(0)
    sh.range('D10').value = QQID
    
    #Write pan list to excel
    sh = xw.books(mqVersion).sheets['CUTLIST']
    for i in range(len(pansRaw)):
        tempRow = 9 + i
        tempNumLoc = 'C' + str(tempRow)
        tempFtLoc = 'D' + str(tempRow)
        tempInLoc = 'E' + str(tempRow)
        sh.range(tempNumLoc).value = pansRaw['Qty'].values[i]
        sh.range(tempFtLoc).value = pansRaw['Feet'].values[i]
        sh.range(tempInLoc).value = pansRaw['In'].values[i]
  
    print("QQtoXL complete") #debug line




# LEAVE QQ to PD as-is 
# Will have to rebuild for XML anyway

def QQtoPD(file):
    print("QQtoPD start") #debug line
    # import data. Rename columns. Count # of rows
    df = pd.read_csv(file, names=['Item', 'Qty', 'Notes', 'Misc'])
    df.drop(df.index[0], inplace=True)
        # import with names list to ensure four columns are initialized
        # drop top row because it is just another label row
    dfCnt = df.shape
    cntRows = dfCnt[0]
    
    #print(df) #debug line
    
    
    #Convert to array to make some things easier...
    dfNp = df.values
    #print(dfNp)
    
    #note locations of sections...
    locSection = []
    for i in range(cntRows):
        tempStr = dfNp[i,0]
        if type(tempStr) != float:
            if 'SECTION' in tempStr:
                locSection.append(i)
    
            
    
    # of sections
    cntSection = len(locSection)
    
    #add cntRows as the "end" of the last section 
    #cntSection does not include this "end" value
    locSection.append(cntRows)
    
    #For loop to reformat CSV from df to dfOut
    
    listOut = []
    panList = []
    
    #Add panel data
    for i in range(cntRows):
        if i in locSection:
            temp = {'Qty': float(df.iloc[i+1,1]), 'Item':df.iloc[i+1,0]}
            listOut.append(temp)
            
    # pull trims, skipping first row of labels and ending prior to first "SECTION" header
    for i in range(cntRows):
        if i > 0 and i < locSection[0]:
            temp = {'Qty': float(df.iloc[i,1]), 'Item': df.iloc[i,0]}
            listOut.append(temp)
    #print(listOut) #debug
    
          
    
    
    #Generate pan list, based on loc
    print(cntSection)
    print(locSection)
    #print(df) #debug
    for i in range(cntSection):
        for j in range(locSection[i]+1, locSection[i+1]):
            str = df.iloc[j,2]
            try:
                locQty = str.index('@')
                tempQty = float(str[:locQty-1])
            except ValueError:
                tempqty = 0
            
            try:
                locFeet = str.index('\'')
                tempFt = int(str[locQty+2:locFeet])
            except ValueError:
                tempFt = 0
            try:
                locIn = str.index('\"')
                tempIn = int(str[locFeet+2:locIn])
            except ValueError:
                tempIn = 0
            tempLen = tempFt * 12 + tempIn
            temp = {'Qty':tempQty, 'Feet':tempFt, 'In':tempIn, 'Length':tempLen}
            panList.append(temp)
    #print(panList)
    
    #convert to df and reorder columns
    dfOut = pd.DataFrame.from_records(listOut)
    dfOut = dfOut[['Qty','Item']]
    
    dfPanList = pd.DataFrame.from_records(panList)
    dfPanList = dfPanList[['Qty', 'Feet','In', 'Length']]
    print(dfPanList)
    
    #Clean any NaN lines from dfOut
    #Some csv input files have a row of ,,,, between the last trim item and SECTION 1
    #Some csv input files do not have the ,,,, row, so itemsConv/dfOut has an extra row of NaN data sometimes
    for i in range(len(dfOut)):  
        if pd.isnull(dfOut.loc[i, 'Qty']):
            dfOut.drop(dfOut.index[i], axis=0, inplace=True)
      
    print("QQtoPD complete") #debug line
    return (dfOut, dfPanList) 


# In[67]:


droppedFile = sys.argv[1]
#droppedFile = "112917-JJ-1CSV16.csv"

QQtoExcel(droppedFile)
input()

