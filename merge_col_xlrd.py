import os
import sys
import re

import xlrd,xlwt
from xlwt import Workbook 

alpha = list('abcdefghijklmnopqrstuvwxyz')

def main(name, r1, r2, col, isRow=False):
    if type(r1) == str:
        isRow = 1
        r1 = alpha.index(r1.lower())
        r2 = alpha.index(r2.lower())
    else:
        r1-=1
        r2-=1
    if type(col) == str:
        col = alpha.index(col.lower())
    else:
        col -= 1
    try:
        rows = []
        for root,dirs,files in os.walk('.'):
            for f in files:        
                if not re.search(r'xlsx?$',f):
                    continue
                p = os.path.abspath( os.path.join(root, f) )
                book = xlrd.open_workbook(p,logfile=open(os.devnull, 'w'))
                extract = re.search(name+r'\(\d+\)\s*(.+?)\(.+xls',p)
                if not extract:
                    continue
                print(f)
                r = [extract.group(1)]
                sht = book.sheet_by_index(0)
                for i in range(r1, r2+1):
                    x = (col, i) if isRow else (i, col)
                    r.append(sht.cell(*x).value)
                rows.append(r)
                #break
        book = Workbook()
        sht = book.add_sheet('2018')
        for r, row in enumerate(rows):
            for c, e in enumerate(row):
                sht.write(r,c,e)
        book.save(rf'C:\Users\xn\Desktop\{name}.xls')
    except Exception as e:
        pass
        # print('warn:'+e)
    finally:
        pass

def check_peixun():
    try:
        rows = []
        for root,dirs,files in os.walk('.'):
            for f in files:        
                if not re.search(r'xlsx?$',f):
                    continue
                if '事业单位工作人员参加培训情况' not in f:
                    continue
                p = os.path.abspath( os.path.join(root, f) )
                book = xlrd.open_workbook(p,logfile=open(os.devnull, 'w'))
                   base = xlrd.open_workbook(p.replace('事业单位工作人员参加培训情况(10)', '聘用合同订立情况(9)'),logfile=open(os.devnull, 'w'))
                px = book.sheet_by_index(0)
                ht = base.sheet_by_index(0)
                d1 = int(ht.cell(8, alpha.index('e')).value or 0)
                d2 = int( px.cell(6, alpha.index('d')).value or 0)
                if d1 != d2:
                    print('管理不一致:'+f,d1,d2)
                if int(ht.cell(22, alpha.index('e')).value or 0) - int(ht.cell(23, alpha.index('e')).value or 0) != int(px.cell(6, alpha.index('l')).value or 0):
                    print('专辑不一致:'+f)
                if int(ht.cell(41, alpha.index('e')).value or 0) - int(ht.cell(42, alpha.index('e')).value or 0) != int(px.cell(6, alpha.index('t')).value or 0):
                    print('工勤不一致:'+f)
    except Exception as e:
        pass
        print('warn:'+str(e))
    finally:
        pass

a = '聘用合同订立情况',8,70,'e'
b = '事业单位公开招聘工作情况',8,67,'f'
c = '事业单位基本情况','i','l',7
check_peixun()
