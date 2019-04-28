import os
import sys
import re
import xlrd,xlwt
from xlwt import Workbook 
import datetime

cd = os.getcwd()
EXCLUDE_REGEX = [
    r'[$~^]',
    r'^汇总数据-'
]
INCLUDE_REGEX = [
]
TITLE_ROW_NUMBER = 1
MIN_VALUES_NUMBER = 2

def main():
    sheets = {}
    for root,dirs,files in os.walk('.'):
        for f in files:        
            if not re.search(r'xlsx?$', f):
                continue
            if len([1 for regex in EXCLUDE_REGEX if re.search(regex, f)]) > 0:
                continue
            if len([1 for regex in INCLUDE_REGEX if not re.search(regex, f)]) > 0:
                continue
            p = os.path.abspath( os.path.join(root, f) )
            book = xlrd.open_workbook(p)
            print(p)
            for name in book.sheet_names():
                if name not in sheets:
                    sheets[name] = []
                sht = book.sheet_by_name(name)
                for r in range(sht.nrows)[TITLE_ROW_NUMBER:]:
                    row = sht.row(r)
                    if len([e.value for e in row if e.value]) < MIN_VALUES_NUMBER :
                        print('    表%s第%s行因为非空白单元格数量少于%s而被忽略:'%(name, r+1, MIN_VALUES_NUMBER), [e.value for e in row])
                        continue
                    sheets[name].append(row)

    book = Workbook()
    for k, rows in sheets.items():
        sht = book.add_sheet(k)
        for r, row in enumerate(rows):
            for c, e in enumerate(row):
                sht.write(r,c, e.value)
    book.save(os.path.join(cd, '汇总数据-%s.xls'%datetime.datetime.now().strftime('%Y%m%d%H%M%S')))

main()
