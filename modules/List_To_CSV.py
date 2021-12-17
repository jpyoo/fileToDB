#!/usr/bin/env python
# coding: utf-8

# # create web category array to import with sql

# In[40]:


#Excel data file, sheet
excelname = 'N41_Website_Category_NInexis_Hidden_Clean.xlsx'
sheetname = 'Hidden_Clean'

#Saving final data as csv
savefilename = 'Hidden.csv'


# In[42]:


import pandas as pd
import json
import numpy as np

def listToCSV(excelname, sheetname, savefilename):
    finalData = pd.DataFrame()
    flist = []

    #Import Website category list json
    #Final category format will be saved on cateogyrList

    cat = pd.json_normalize(pd.read_json('webcategory.json')['children_data'])
    cat2 = pd.DataFrame(columns = cat.columns)

    for i in range(len(cat)):
        cat2 = cat2.append(cat['children_data'][i])

    categoryList = cat.append(cat2).sort_values(by=['id'])

    #read Excel file and split category list
    #category list only contains id
    #convert the ids to publishable format
    #save style# and converted category string as CSV
    
    xls = pd.read_excel(excelname, sheet_name=sheetname)
    xls['Split Category'] = xls['N41 Category'].str.split(',')
    xls.dropna(subset = ['N41 Category'], inplace = True)
    flist = []
    for idList in xls['Split Category']:
        newid = []
        for it in idList:
            try:
                newid.append("[" + str(it) + "] " + categoryList[categoryList['id'] == int(it)]['name'].iloc[0])

            except:
                pass
        flist.append(','.join(newid)) 

    finalData['style'] = pd.Series(xls['styleNo'])
    finalData['attributeValue'] = flist

    finalData.to_csv(savefilename)

