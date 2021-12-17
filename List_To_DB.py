#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from modules.DBsql import N41DB
from modules.List_To_CSV import listToCSV


# In[5]:


import json

filedata = json.load(open('excelName.json'))
#Excel data file, sheet
excelname = filedata['excelname']
sheetname = filedata['sheetname']

#Saving final data as csv
savefilename = filedata['savefilename']
siteid = filedata['siteid']


# In[ ]:


n41db  = N41DB()
n41db.loadKey('db.key')

listToCSV(excelname, sheetname, savefilename)
n41db.CSVtoDB_styleAttribute(siteid)
n41db.CSVtoDB_sitePublish(siteid)

