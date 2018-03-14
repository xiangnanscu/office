import os
import sys
import re
import shutil
import comtypes.client
import glob


cd = os.getcwd()
word = None
excel = None

def out(s):
    if word:
        word.Quit()
    if excel:
        excel.Quit()
    return exit(s)

def _(s):
    return os.path.join(cd, s)

def clean(s):
    return re.sub(r'\s', '', s)

def delBlank(s):
    return re.sub(r'\s','',s)

books = glob.glob('*.xls') + glob.glob('*.xlsx') 
books = [e for e in books if not('~' in e or '$' in e)]
if not books:
    out('没有找到数据文件')
excel_name = books[0]
print('找到数据文件：'+excel_name)
docs = glob.glob('*.doc') + glob.glob('*.docx') 
docs = [e for e in docs if not('~' in e or '$' in e)]
if not docs:
    out('没有找到word文件')

try:
    excel = comtypes.client.CreateObject('Excel.Application')
except Exception as e:
    out('Excel程序启动失败：'+str(e))
try:
    book = excel.Workbooks.Open(_(excel_name))
except Exception as e:
    out('打开数据文件失败：'+str(e))

UsedRange = book.ActiveSheet.UsedRange
sht = book.ActiveSheet
data = {}    
for r in range(2, UsedRange.Rows.Count+1):
    data[sht.Cells(r, 1).Text] = sht.Cells(r, 2).Text
book.Close()

try:
    word = comtypes.client.CreateObject('Word.Application')
except Exception as e:
    out('WORD程序启动失败：'+str(e))


def make_doc(doc_name, d):
    try:
        doc = word.Documents.Open(_(doc_name))
    except Exception as e:
        return '打开模板文件失败：'+str(e) 
    try:
        r = word.Selection
        for old, new in d.items():
            print('  已将%s替换为%s'%(old, new))
            r.Find.Text = old
            r.Find.Replacement.Text = new
            r.Find.Execute(Replace=2)
        doc.Save()      
    except Exception as e:
        return '替换文件失败：'+str(e) 

for doc_name in docs:
    print(doc_name)
    err = make_doc(doc_name, data)
    if err:
        print(err)

out('==========程序结束===============')



