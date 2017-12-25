import os
import sys
import re
import shutil
import comtypes.client

wenHaoPat = r'\s*(\w+)\s*〔([\s\d]+)〕([\d\s-]+)号'

wdFormatPDF = 17
cd = os.getcwd()

try:
    word = comtypes.client.CreateObject('Word.Application')
except Exception as e:
    print('创建word程序失败：'+str(e))
    exit()

def delBlank(s):
    return re.sub(r'\s','',s)

for root, dirs, files in os.walk('.'):
    for f in files:
        if '~' in f or not re.search(r'docx?$',f):
            continue
        docpath = os.path.abspath( os.path.join(root, f) )
        try:
            doc = word.Documents.Open(docpath)
        except Exception as e:
            print('打开文档'+docpath+'失败：'+str(e))
            continue
        print('处理:'+docpath)
        # if len(doc.Paragraphs) > 50:
        #     continue
        Paragraphs = [p.Range for p in doc.Paragraphs]
        ext = f.split('.')[-1] or ''
        newname = ''
        title = ''
        for r in Paragraphs:
            if r.Font.Size == 22: # r.Font.Name == "方正小标宋简体" and 
                title += r.Text
        title = re.sub(r'\s|(江安县职称改革工作领导小组办公室)|(江安县人力资源和社会保障局)','',title)
        for r in Paragraphs:
            m = re.search(wenHaoPat, r.Text)
            if m:
                newpath = os.path.join(cd, delBlank(m.group(1)))
                year = delBlank(m.group(2))
                number = delBlank(m.group(3))
                if number and not os.path.exists(newpath):
                    os.makedirs(newpath)
                newname = os.path.join(newpath, '%s-%s号-%s.%s'%(year,number,title,ext)) 
                break
        else:
            print('  没有找到文号')
        doc.Close()
        if newname:
            if not os.path.exists(newname):  
                try:
                    shutil.move(docpath, newname)
                    print('  移动: %s-->%s'%(docpath, newname))
                except Exception as e:
                    print('  警告：%s移动失败:%s'%(newname, e))
            else:
                print('  警告：%s已存在，移动文件失败'%(newname,))

word.Quit()
 


