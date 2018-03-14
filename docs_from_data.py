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
    out('没有找到模板文件')
doc_name = docs[0]
print('找到模板文件：'+doc_name)

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
data = []    
for r in range(1, UsedRange.Rows.Count):
    d = {}
    if sht.Cells(r+1, 1).Text == '':
        break
    for c in range(UsedRange.Columns.Count):
        d[sht.Cells(1, c+1).Text] = clean(sht.Cells(r+1, c+1).Text)
    data.append(d)
book.Close()

try:
    word = comtypes.client.CreateObject('Word.Application')
except Exception as e:
    out('WORD程序启动失败：'+str(e))


def make_doc(d, name):
    print('生成:'+name)
    print(' ', d)
    try:
        doc = word.Documents.Open(_(doc_name))
    except Exception as e:
        return '打开模板文件失败：'+str(e) 
    try:
        f = word.Selection.Find
        for k, v in d.items():
            f.Text = '{%s}' % k
            f.Replacement.Text = v
            f.Execute(Replace=2)
        doc.SaveAs(name)       
    except Exception as e:
        return '生成文件失败：'+str(e) 

res_dir = os.path.join(cd, '结果')
if not os.path.exists(res_dir):
    os.makedirs(res_dir)
for i, d in enumerate(data):
    name = '%s-%s'%(i, doc_name)
    err = make_doc(d, os.path.join(res_dir,name))
    if err:
        out(err)

out('==========程序结束===============')



