import os
import sys
import re

import xlrd,xlwt
from xlwt import Workbook 


try:
    rows = []
    for root,dirs,files in os.walk('.'):
        for f in files:        
            if not re.search(r'xlsx?$',f):
                continue

            p = os.path.abspath( os.path.join(root, f) )
            
            book = xlrd.open_workbook(p)
            extract = re.search(r'事业单位基本情况\(1\)\s*(.+?)\(\d+\) \.xls',p)
            if not extract:
                continue
            r = [extract.group(1)]
            sht = book.sheet_by_index(0)

            for i in range(5, 31):
                r.append(sht.cell(i, 4).value)
            rows.append(r)
            #break
    book = Workbook()
    sht = book.add_sheet('2016')
    for r, row in enumerate(rows):
        for c, e in enumerate(row):
            sht.write(r,c,e)
    book.save(r'C:\Users\xn\Desktop\total.xls')
except Exception as e:
    print('warn:'+e)
finally:
    pass
 

