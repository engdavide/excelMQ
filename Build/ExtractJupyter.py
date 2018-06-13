
# coding: utf-8

# In[24]:


import pandas as pd
import xlwings as xw
import datetime
import os
import sys

now = datetime.datetime.now()
now = str(now.year)+'-'+str(now.month)+'-'+str(now.day)+'-'+str(now.hour)+str(now.minute)


# In[33]:


# --Pull savable data from Excel
def MQtoCSV():
    mqVersion = 'MultiQuoter2018v0' #This is the filename for current MQ
    #designate sheet to pull from
    sh = xw.books(mqVersion).sheets['MULTIQUOTER']
    
    #Pull header data in format of Field: Value:
    headerKeys = ['Customer','Substrate','Paint','Color','RibH','PanW','QQID',
              'PriceOR','Tax','Delivery','Rep']
    headerLocs = ['C3','C5','C6','C7','C8','C9','D10','H2','H3','H4','H5']
    headerVals = []
    for i in range(len(headerLocs)):
        temp = sh.range(headerLocs[i]).value
        headerVals.append(temp)
    headerVals.append(now)
    headerKeys.append('timeStamp')
    dfHeader = pd.DataFrame({'Field': headerKeys,
                             'Value' : headerVals})
    print(dfHeader)
   
    #--Pull item list data in format of line: Qty, Code, Comments, etc...
    
    #--non-pan List:
    #Count num of active rows
    numRows = 30
    for i in range(numRows):
        loc = ('B' + str(i + 17))
        if pd.isnull(sh.range(loc).value):
            numRows = i 
            break
    #initialize temp columns...
    itemQty = []
    itemCode = []
    itemComment = []
    itemPriceOR = []
    itemColorOR = []
    itemMetOR = []
    #Run FOR loops to build each column...
    for i in range(numRows):
        loc = ('B' + str(i + 17))
        itemQty.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('C' + str(i + 17))
        itemCode.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('K' + str(i + 17))
        itemComment.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('M' + str(i + 17))
        itemPriceOR.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('O' + str(i + 17))
        itemColorOR.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('Q' + str(i + 17))
        itemMetOR.append(sh.range(loc).value)
    #Build the DF
    dfItems = pd.DataFrame({'itemQty': itemQty,
                            'itemCode' : itemCode,
                            'itemComment' : itemComment,
                            'itemPriceOR' : itemPriceOR,
                            'itemColorOR' : itemColorOR,
                            'itemMetOR' : itemMetOR})
    #Reorder
    dfItems = dfItems[['itemQty', 'itemCode', 'itemComment', 'itemPriceOR',
                       'itemColorOR', 'itemMetOR']]
    print(dfItems)

    
   #--PAN List:
    #Count num of active rows
    numRows = 1
    for i in range(numRows):
        loc = ('B' + str(i + 14))
        if pd.isnull(sh.range(loc).value):
            numRows = i 
            break
    #initialize temp columns...
    panQty = []
    panCode = []
    panComment = []
    panPriceOR = []
    panColorOR = []   
    panPanOR = []
    panMetOR = []
    #Run FOR loops to build each column...
    for i in range(numRows):
        loc = ('B' + str(i + 14))
        panQty.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('C' + str(i + 14))
        panCode.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('K' + str(i + 14))
        panComment.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('M' + str(i + 14))
        panPriceOR.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('O' + str(i + 14))
        panColorOR.append(sh.range(loc).value)
    for i in range(numRows):
        loc = ('Q' + str(i + 14))
        panMetOR.append(sh.range(loc).value)
    #Build the DF
    dfPans = pd.DataFrame({'panQty': panQty,
                           'panCode' : panCode,
                           'panComment' : panComment,
                           'panPriceOR' : panPriceOR,
                           'panColorOR' : panColorOR,
                           'panMetOR' : panMetOR})
    #Reorder
    dfPans = dfPans[['panQty', 'panCode', 'panComment', 'panPriceOR',
                     'panColorOR', 'panMetOR']]
    print(dfPans)


    #CUTLIST from CUTLIST sheet
    #Designate appropriate sheet...
    shCut = xw.books(mqVersion).sheets['CUTLIST']
    #Everything will be run in an A and B list
    # A list for first panel data, B list for 2nd
    
    #Count num of active rows
    numRowsA = 200
    numRowsB = 200
    for i in range(numRows):
        loc = ('C' + str(i + 9))
        if pd.isnull(shCut.range(loc).value):
            numRowsA = i 
            break
    for i in range(numRows):
        loc = ('I' + str(i + 9))
        if pd.isnull(shCut.range(loc).value):
            numRowsB = i 
            break
    print(numRowsA)
    print(numRowsB)
    #NOTE for some reason, this is not working...
        
    #initialize temp columns...
    cutNumA = []
    cutFtA = []
    cutInA = []
    cutNumB = []
    cutFtB = []
    cutInB = []
    #Run FOR loops to build each column...
    for i in range(numRowsA):
        loc = ('C' + str(i + 9))
        cutNumA.append(shCut.range(loc).value)
    for i in range(numRowsA):
        loc = ('D' + str(i + 9))
        cutFtA.append(shCut.range(loc).value)
    for i in range(numRowsA):
        loc = ('E' + str(i + 9))
        cutInA.append(shCut.range(loc).value)
    #FOR loops for list B
    for i in range(numRowsB):
        loc = ('I' + str(i + 9))
        cutNumB.append(shCut.range(loc).value)
    for i in range(numRowsB):
        loc = ('J' + str(i + 9))
        cutFtB.append(shCut.range(loc).value)
    for i in range(numRowsB):
        loc = ('K' + str(i + 9))
        cutInB.append(shCut.range(loc).value)
    #Build the DF. Must use separate DF's for each A and B
    #Otherwise, python doesn't seem to like arrays of differing lengths
    #probably bad practice, but changing the col names to cutANum to sort
        #out better in the csv, but leaving naming convention otherwise unchanged
    dfCutA = pd.DataFrame({'cutANum': cutNumA,
                           'cutAFt' : cutFtA,
                           'cutAIn' : cutInA})
    dfCutB = pd.DataFrame({'cutBNum': cutNumB,
                           'cutBFt' : cutFtB,
                           'cutBIn' : cutInB})
    #Reorder
    dfCutA = dfCutA[['cutANum','cutAFt','cutAIn']]
    dfCutB = dfCutB[['cutBNum','cutBFt','cutBIn']]
    print(dfCutA)
    print(dfCutB)    

    
    #--write to csv--
    
    #Set filename
    cust = sh.range('C2').value
    address = sh.range('C11').value
    filename = str(cust)+' - '+str(address)+' - '+str(now)+'.csv'
    path = os.path.dirname(sys.argv[0])
    path = os.path.dirname(path)
    filePath = path + '\\Quotes\\' + filename
    print(filePath)
    
    #try to concat DF's together...
    dfAll = pd.concat([dfHeader, dfPans, dfItems, dfCutA, dfCutB])

    #to csv
    dfAll.to_csv(filePath, encoding='utf-8', index=False)
    


# In[34]:


MQtoCSV()
input()
