# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 16:13:59 2024

@author: Five-seveN
"""

import glob
import os
import pandas as pd


s_sheet = [] #s_sheet is the sheet that will be used to do main calculation and has pretty look
sheetname = ['1-10月份請款單 A','1-10月份請款單B','1-10月份請款單CD','1-10月份請款單EFG','1-10月份請款單HI','1-10月份請款單JKLM','1-10月份請款單NOPQ','1-10月份請款單RSTU','1-10月份請款單X','1-10月份請款單YZ']
for order in sheetname:
    str_form = pd.read_excel('testwrite.xlsx', order)
    s_sheet.append(pd.DataFrame.to_numpy(str_form)) #complex double list for include all sheets at the same structure

complist = []
location = []
buffer = []

#fetch all companies' name and make a list of them in order of s_sheet
for k in range(len(s_sheet)):
    sublocation = []
    subcomplist = []
    # These two sublist is for matching the inner structure of the company name and location to the s_sheet's structure
    for n in range(len(s_sheet[k])):
        if s_sheet[k][n,0] == s_sheet[k][0,0] and pd.isna(s_sheet[k][n,1]) == False:
        #The seond is for making sure there will be no empty entry that fail all shit
        #Notice that first criteria requires the r1c1 position be "TO:" or it'll break, this might cause future troubles
            buffer = s_sheet[k][n,1]
            buffer = buffer.replace(" ","")
            subcomplist.append(buffer)
            sublocation.append(n)
        else:
            continue
    location.append(sublocation)
    complist.append(subcomplist)
    
import xlwings as xw
wb = xw.Book('testwrite.xlsx')

import openpyxl
paidlist = []

# Now we have all location of the big sheet again, we shall read the marks X that noted in the files.
for k in range(len(complist)):
    for n in range(len(complist[k])):
        if wb.sheets(sheetname[k]).range('E'+ str(location[k][n]+10)).value == "X":
            print(location[k][n])
            print(complist[k][n])
            paidlist.append(complist[k][n]) #the output is a paidlist including all the companies that had paid.

del(str_form,buffer,complist,location,s_sheet,order,wb,k,n)

#the below code write the paid status back to 對照表

wb = xw.Book('對照表.xlsx') #not yet sure how mom's new matching sheet is, so this must do on site adjustment of file type
str_form = pd.read_excel('對照表.xlsx')
sheet = pd.DataFrame.to_numpy(str_form)
for clean in range(len(sheet[:,1])):
    sheet[clean,1] = sheet[clean,1].replace(" ","") #cleaning up all the annoying bubbles

for k in paidlist:
    for n in range(len(sheet[:,1])): #this, again, not neccesarily first column, so need adjustments.
        if sheet[n,1] == k:
            wb.sheets('Sheet4').range('G' + str(n + 2)).value = 'X'