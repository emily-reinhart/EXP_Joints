#!/usr/bin/env python
# coding: utf-8

# In[6]:


import xlrd
from decimal import Decimal
import pandas as pd


# In[29]:


# load excel file with data
# loc = ".\\data.xlsm"
loc = ".\\" + input("Please enter file name with file type extension: ")

wb = xlrd.open_workbook(loc)
# sheet = wb.sheet_by_index(0)


# In[27]:


# ask user for pairs they want to search
pairsInput = input('Type in pairs to evaluate, separate each number with a comma. Ex. 120,130,140,150')

inputList = pairsInput.split(',')
pairsList = []

for x in range(0,len(inputList)-1,2):
  pairsList.append([inputList[x], inputList[x+1]])


# In[30]:


# print(sheet.cell_value(6,0), sheet.cell_value(6,1), sheet.cell_value(6,2), sheet.cell_value(6,3), sheet.cell_value(6,4), sheet.cell_value(6,5), sheet.cell_value(6,6))

newTable = []

tableColumns = sheet.row_values(6,0,7)

for pair in pairsList:
    for sheet in wb.sheets():
        # print(sheet.row_values(6,0,7))
        for row_num in range(9,sheet.nrows,1):
                row_value = sheet.row_values(row_num,0,7)
                if row_value[0] == int(pair[0]):
                    #print(row_value)
                    row1 = row_value
                if row_value[0] == int(pair[1]):
                    #print(row_value)
                    row2 = row_value

        # print(row1, row1[4], row2[4], ((row1[4]) - (row2[4])))
        separator = ' - '
        newRow = [separator.join(pair),(row1[1]-row2[1]),(row1[2]-row2[2]),(row1[3]-row2[3]),(row1[4]-row2[4]),(row1[5]-row2[5]),(row1[6]-row2[6])]
        # print(newRow)
        newTable.append(newRow)
    
pd.DataFrame(newTable, columns=tableColumns)
