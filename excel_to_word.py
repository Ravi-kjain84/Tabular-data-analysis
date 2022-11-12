# -*- coding: utf-8 -*-
"""
Created on Sat Nov 12 13:39:46 2022

@author: DELL
"""

import datetime as dt
import random
import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate



df = pd.read_excel(r'C:\Users\DELL\Desktop\Python practice\Word Template\reports\input.xlsx',\
                   sheet_name = 'Sheet1',skiprows=1)

new_salesTblRows = []    
for index,row in df.iterrows():
    new_salesTblRows.append({"sNo": df.loc[index,'sNo'],
                         "name": df.loc[index,'name'],
                         "cPu": df.loc[index,'cPu'],
                         "nUnits": df.loc[index,'nUnits'],
                         "revenue": df.loc[index,'revenue'],
                         "item": df.loc[index,'item']})
    
doc = DocxTemplate("reportTmpl01.docx")

# create data for reports
salesTblRows = []
for k in range(10):
    costPu = random.randint(1, 15)
    nUnits = random.randint(100, 500)
    salesTblRows.append({"sNo": k+1, "name": "Item "+str(k+1),
                         "cPu": costPu, "nUnits": nUnits, "revenue": costPu*nUnits})

topItems = [x["name"] for x in sorted(salesTblRows, 
                                      key=lambda x: x["revenue"], 
                                      reverse=True)][0:3]

todayStr = dt.datetime.now().strftime("%d-%b-%Y")


# create context to pass data to template
context = {
    "reportDtStr": todayStr,
    "salesTblRows": new_salesTblRows,
    "topItemsRows": topItems
}

# render context into the document object
doc.render(context)

# save the document object as a word file
reportWordPath = 'reports/report_{0}.docx'.format(todayStr)
doc.save(reportWordPath)

# convert the word file as pdf file
convert(reportWordPath, reportWordPath.replace(".docx", ".pdf"))
