import os
import sys
import re
import collections

import xlrd

join = os.path.join

p = join(os.getcwd(), 'aa.xls')


ret = collections.OrderedDict()
data = xlrd.open_workbook(p)


table = data.sheets()[0] 



    

for c in range(table.ncols):
    name = table.cell(0, c).value
    if re.match(r'^[^a-z]+$', name):
        ret[name] = s = set()
        for r in range(1, table.nrows):
            s.add(table.cell(r, c).value)

for k, v in ret.items():
    if len(v)<100:
        print(k,v)
