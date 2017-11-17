import os
import sys
import re

import comtypes.client

wdFormatPDF = 17


app = comtypes.client.CreateObject('Excel.Application')

sheets = {}
for root,dirs,files in os.walk('.'):
    for f in files:        
        if not re.search(r'xlsx?$',f):
            continue

        p = os.path.abspath( os.path.join(root, f) )
        book = app.Workbooks.open(p)
        for sht in book.Worksheets:
        	if sht.Name not in sheets:
        		sheets[sht.Name] = []
        	print(sht.UsedRange.Rows.Count)
        	for i in range(sht.UsedRange.Rows.Count):
        		row = []
        		if i < 3:
        			continue
        		c = sht.UsedRange.Columns.Count
        		for j in range(c):
        			if not sht.Cells[i+1, j+1].Text:
        				continue
		        	row.append(sht.Cells[i+1, j+1].Text)
		        if len(row) > 6:
		        	sheets[sht.Name].append(row)
        #break

book = app.Workbooks.Add()
for k, v in sheets.items():
	t = 0
	sht = book.Worksheets.add()
	sht.Name = k
	for r, row in enumerate(v):
		for c, e in enumerate(row):
			sht.Cells[t+1, c+1] = e
		t += 1
book.SaveAs(r'C:\Users\Administrator\Desktop\total.xlsx')

app.Quit()
 


