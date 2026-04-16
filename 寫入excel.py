
# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 03:07:11 2024

@author: Five-seveN
"""

filename = '115年1-10月份請款表.xlsx'

import glob
import os
import pytesseract
from PIL import Image
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# Notice one have to install tesseract and locate its position first

imgtable = []

#fetch image files
path = os.path.split((os.path.abspath(__file__)))[0]
os.chdir(path)
for file in glob.glob("*.png"):
    imgtable.append(file)
    
money = []
Unicode = [] #really important list, this list has its order aligned with the money
companies = [] #would be used later on for matching UniCode to companies

for names in imgtable:
    # Open the image file
    image = Image.open(names).convert('L')
    
    #define cropping area and crop it
    box = (3060,730,3220,780)
    crpimg = image.crop(box)
    
    # Perform OCR using PyTesseract, use buffer to solve and replace the unwanted new line
    buffer = pytesseract.image_to_string(crpimg)
    buffer = buffer.replace("\n","")
    buffer = buffer.replace(",","")
    if buffer == '':
        buffer = '0'
    money.append(buffer)
    Unicode.append(int(names[0:8]))

# Print the extracted text
print(money)
print(Unicode)
del(file,imgtable,image,box,buffer,names,crpimg)



# The second part we'll read companies' relative name by matching their UniCode
import pandas as pd
m_sheet_str = pd.read_excel('對照表.xlsx')
m_sheet = pd.DataFrame.to_numpy(m_sheet_str)

#the below code is just for unifying the data type of unicode we uses and removing all weird bubbles of spaces
k = 0
for clear in m_sheet[:,2]:
    m_sheet[k,1] = m_sheet[k,1].replace(" ","")
    # removing bubbles from companies' name
    if type(clear) == int:
        k = k+1
        continue
    # Skip all that is already int so "replace" function doesn't go error
    m_sheet[k,2] = m_sheet[k,2].replace(" ","")
    if m_sheet[k,2] == '':
        m_sheet[k,2] = 0
    # This if is for preventing error with lines that has no entry, since WE ARE NOT SURE HOW MANY SPACES IT READ
    else:
        m_sheet[k,2] = int(m_sheet[k,2])
    k = k+1

# below are for matching codes with companies' name
for code in Unicode:
    k = 0
    if all(match != code for match in m_sheet[:,2]):
        companies.append(code)
    for match in m_sheet[:,2]:
        m_sheet[k,2] = m_sheet[k,2]
        if match == code:
            companies.append(m_sheet[k,1])
        k = k+1

print(companies)
del(k,clear,code,match,m_sheet_str)



#The third part we'll call in the sheet that are mainly for calculating the data
s_sheet = [] #s_sheet is the sheet that will be used to do main calculation and has pretty look
sheetname = ['1-10月份請款單 A','1-10月份請款單B','1-10月份請款單CD','1-10月份請款單EFG','1-10月份請款單HI','1-10月份請款單JKLM','1-10月份請款單NOPQ','1-10月份請款單RSTU','1-10月份請款單X','1-10月份請款單YZ']
for order in sheetname:
    str_form = pd.read_excel(filename, order)
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
wb = xw.Book(filename)

# Now we have all location in a fine list, we need to take them as reference and replace them one by one.
for m in range(len(companies)):
    for k in range(len(complist)):
        # These two sublist is for matching the inner structure of the company name and location to the s_sheet's structure
        for n in range(len(complist[k])):
            if companies[m] == complist[k][n]:
                print(location[k][n])
                wb.sheets(sheetname[k]).range('C'+ str(location[k][n]+4)).value = money[m]