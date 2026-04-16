# -*- coding: utf-8 -*-
"""
Created on Mon Jan  8 23:30:56 2024

@author: Five-seveN
"""

# importing all the required modules
import pypdf
from pypdf import PdfReader
import glob
import os

#construct table for companies, files
table = []

#fetch all files and make tables
os.chdir(r"C:\Users\lenovo\Desktop\113年1-10月請款表")
for file in glob.glob("*.pdf"):
    table.append(file)
    
for names in table:
    # creating a pdf reader object
    reader = PdfReader(names)
    
    # print the number of pages in pdf file
    print(len(reader.pages))
    
    # extract images
    page = reader.pages[0]
    with open(names[6:14] + '.png', "wb") as fp: # because naming system of files, 6 to 14 happens fetch 統一編號
        fp.write(page.images[0].data) #this can fetch image data and write them into png files