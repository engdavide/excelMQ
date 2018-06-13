
# coding: utf-8

# In[12]:


import pandas as pd
import xlwings as xw
import sys


# In[49]:


def CSVtoExcel(file):    

    # import data
    dfIn = pd.read_csv(file)

    #--Split into 5 DF's, reorder, drop empty rows from pan and item lists, replace NaN's with "" in 
    #item and pan lists, reset indices--
    dfHeader = dfIn.loc[:,'Field':'Value']
    dfHeader = dfHeader.dropna(axis=0, how='all')
    dfHeader.reset_index(drop=True, inplace=True)
    
    dfItems = dfIn.loc[:,'itemCode' : 'itemQty']
    dfItems = dfItems[['itemQty', 'itemCode', 'itemComment', 'itemPriceOR',
                       'itemColorOR', 'itemMetOR']]
    dfItems = dfItems.dropna(axis=0, how='all')
    dfItems.reset_index(drop=True, inplace=True)
    dfItems = dfItems.fillna('')
    
    dfPans = dfIn.loc[:,'panCode' : 'panQty']
    dfPans = dfPans[['panQty', 'panCode', 'panComment', 'panPriceOR',
                     'panColorOR', 'panMetOR']]
    dfPans = dfPans.dropna(axis=0, how='all')
    dfPans.reset_index(drop=True, inplace=True)
    dfPans = dfPans.fillna('')

    dfCutA = dfIn.loc[:,'cutAFt':'cutANum']
    dfCutA = dfCutA[['cutANum','cutAFt','cutAIn']]
    dfCutA = dfCutA.dropna(axis=0, how='all')
    dfCutA.reset_index(drop=True, inplace=True)

    dfCutB = dfIn.loc[:,'cutBFt':'cutBNum']
    dfCutB = dfCutB[['cutBNum','cutBFt','cutBIn']]
    dfCutB = dfCutB.dropna(axis=0, how='all')
    dfCutB.reset_index(drop=True, inplace=True)


    
    print(dfHeader)
    print(dfItems)
    print(dfPans)
    print(dfCutA)
    print(dfCutB)
    
    #--Write to Excel--
    mqVersion = 'MultiQuoter2018v0' #This is the filename for current MQ
    #designate sheet to write to
    sh = xw.books(mqVersion).sheets['MULTIQUOTER']
    
    #Header
    headerLocs = ['C3','C5','C6','C7','C8','C9','D10','H2','H3','H4','H5']
    for i in range(len(headerLocs)):
        #print(headerLocs[i])
        #print(dfHeader.at[i,'Value'])
        sh.range(headerLocs[i]).value = dfHeader.at[i,'Value']
    
    #Pans
    numRows = len(dfPans)

    # FOR loops to build each column...
    for i in range(numRows):
        loc = ('B' + str(i + 14))
        #print(loc)
        #print(i)
        #print(dfPans.at[i,'panQty'])
        sh.range(loc).value = dfPans.at[i,'panQty']
    for i in range(numRows):
        loc = ('C' + str(i + 14))
        sh.range(loc).value = dfPans.at[i,'panCode']
    for i in range(numRows):
        loc = ('K' + str(i + 14))
        sh.range(loc).value = dfPans.at[i,'panComment']
    for i in range(numRows):
        loc = ('M' + str(i + 14))
        sh.range(loc).value = dfPans.at[i,'panPriceOR']
    for i in range(numRows):
        loc = ('O' + str(i + 14))
        sh.range(loc).value = dfPans.at[i,'panColorOR']
    for i in range(numRows):
        loc = ('Q' + str(i + 14))
        sh.range(loc).value = dfPans.at[i,'panMetOR']

        
    #Items
    numRows = len(dfItems)

    # FOR loops to build each column...
    for i in range(numRows):
        loc = ('B' + str(i + 17))
        #print(loc)
        #print(i)
        #print(dfPans.at[i,'panQty'])
        sh.range(loc).value = dfItems.at[i,'itemQty']
    for i in range(numRows):
        loc = ('C' + str(i + 17))
        sh.range(loc).value = dfItems.at[i,'itemCode']
    for i in range(numRows):
        loc = ('K' + str(i + 17))
        sh.range(loc).value = dfItems.at[i,'itemComment']
    for i in range(numRows):
        loc = ('M' + str(i + 17))
        sh.range(loc).value = dfItems.at[i,'itemPriceOR']
    for i in range(numRows):
        loc = ('O' + str(i + 17))
        sh.range(loc).value = dfItems.at[i,'itemColorOR']
    for i in range(numRows):
        loc = ('Q' + str(i + 17))
        sh.range(loc).value = dfItems.at[i,'itemMetOR']



    #Write cut list
    shCut = xw.books(mqVersion).sheets['CUTLIST']

    numRowsA = len(dfCutA)
    numRowsB = len(dfCutB)

    for i in range(numRowsA):
        loc = ('C' + str(i+9))
        shCut.range(loc).value = dfCutA.at[i,'cutANum']
    for i in range(numRowsA):
        loc = ('D' + str(i+9))
        shCut.range(loc).value = dfCutA.at[i,'cutAFt']
    for i in range(numRowsA):
        loc = ('E' + str(i+9))
        shCut.range(loc).value = dfCutA.at[i,'cutAIn']

# In[50]:


droppedFile = sys.argv[1]
#droppedFile = "KROCZYNSKY ROBERT - 1123 some sstree - 2017-12-1-932.csv"

CSVtoExcel(droppedFile)
input()

