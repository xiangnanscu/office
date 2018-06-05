var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
var yargs = require('yargs');
var xlsx = require('xlsx');

function sameObject(a, b) {
  let k1 = Object.keys(a)
  let k2 = Object.keys(b)
  if (k1.length !== k2.length) {
    return false
  }
  for (let k of k1) {
    if (a[k] !== b[k]) {
        return false
    }
  }

  return true
}

function getSheets(fn) {
    var workbook = xlsx.readFile(fn, {
      type: 'binary'
    });
    let sheets = {}
    for (let name in workbook.Sheets) {
      let sht = workbook.Sheets[name]
      sheets[name] = xlsx.utils.sheet_to_json(sht)
    }
    return sheets
}

function hasSameRow(r, rows) {
    for (let row of rows) {
        if (sameObject(r, row)) {
            return true
        }
    }
}

function template(templateName, outputName, data) {
    console.log(data)
  var zip = new JSZip(fs.readFileSync(templateName, 'binary'));
  var doc = new Docxtemplater();
  doc.loadZip(zip);
  //set the templateVariables
  doc.setData(data||{});
  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
  }
  catch (error) {
    var e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
    }
    console.log(JSON.stringify({error: e}));
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
    throw error;
  }
  var buf = doc.getZip().generate({type: 'nodebuffer'});
  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  fs.writeFileSync(outputName, buf);
}

function getValue(list, key, Factory) {
    let value = list.get(key)
    if (value===undefined) {
        value = new Factory()
        list.set(key, value)
    }
    return value
}

class Base {
    constructor(data, name) {
        this.data = data
        this.name = name
        this.now = new Date()
    }
    cleanData() {
        let d = this.data
        let now = this.now
        d['下属单位'] = d['下属单位'] || ''
        d['文号年份'] = d['文号年份'] || now.getFullYear().toString()
        d['成文日期'] = d['成文日期'] || `${now.getFullYear()}年${now.getMonth()+1}月${now.getDate()}日`
         if (d['来函文号']) {
             d['来函文号'] = d['来函文号'].replace(/[【[]/,'〔').replace(/[\]】]/,'〕')
         }    
    }
    getFileNamePrefix() {
        return `${this.name}/${this.data['文号年份']}-${this.data['文号']}-`
    }
    getFileName() {

    }
    getTemplateName() {
        let suffix = this.data.S ? 'S' : ''
        return `模板/${this.name}${suffix}.docx`
    }
    exec() {
        this.cleanData()
        let savePath = this.getFileNamePrefix() + this.getFileName()
        if (fs.existsSync(savePath)) {
            return console.log('文件存在，制作失败：'+savePath)
        }
        template(
            this.getTemplateName(), 
            savePath, 
            this.data)
        console.log('制作成功：'+savePath)
    }
}
class 辞聘 extends Base {
    cleanData() {
        super.cleanData()
    }
    getFileName() {
        return `关于${this.data['姓名']}同志辞聘备案的函.docx`
    }
}
class 岗位聘任 extends Base {
    static mergeData(rows) {
        let first = rows[0]
        let items = new Map()
        for (let data of rows) {
            let e = getValue(items, data['下属单位'], Map)
            e = getValue(e, data['岗位'], Array)
            e.push(data['姓名'])
        }
        let body = [...items].map(([subName, works])=>{
            let infos = [...works].map(([work, names])=>{
                if (work.slice(-2) !== '岗位') {
                    work = work + '岗位'
                }
                return `聘任${names.join('、')}同志到${work}`
            })
            return `${subName}${infos.join('，')}`
        })
        let titleNames
        if (rows.length==2) {
            titleNames = `${rows[0]['姓名']}${rows[1]['姓名']}二名`
        } else {
            titleNames = `${first['姓名']}等${rows.length}名`
        }
        return {
            ...first, 
            S: true,
            '姓名': titleNames, 
            '聘任明细': body.join('；'),
        }
    }
    cleanData() {
        super.cleanData()
        let data = this.data
        data['任职时间'] = data['任职时间'] || data['成文日期']
    }
    getFileName() {
        return `关于${this.data['姓名']}同志岗位聘任备案的函.docx`
    }
}
class 招用 extends Base {
    static mergeData(rows) {
        let first = rows[0]
        let items = new Map()
        for (let data of rows) {
            let e = getValue(items, data['下属单位'], Array)
            e.push(data['姓名'])
        }
        let body = [...items].map(([subName, names])=>{
            return `${subName}招用${names.join('、')}同志为劳动合同制工作人员`
        })
        let titleNames
        if (rows.length==2) {
            titleNames = `${rows[0]['姓名']}${rows[1]['姓名']}二名`
        } else {
            titleNames = `${first['姓名']}等${rows.length}名`
        }
        return {
            ...first, 
            S: true,
            '姓名': titleNames, 
            '招用明细': body.join('；'),
        }
    }
    cleanData() {
        super.cleanData()
    }
    getFileName() {
        return `关于确认招用${this.data['姓名']}同志为劳动合同制工作人员的函.docx`
    }
}

function mergeRows(rows, key) {
    let merged = new Map()
    for (let data of rows) {
        getValue(merged, data[key], Array).push(data)
    }
    return merged
}

function inputExit(s) {
    const readline = require('readline');
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    rl.question(s, (answer) => {
      rl.close();
    });  
}

try {
    let last = getSheets('./data.last.xlsx')
    let current = getSheets('./data.xlsx')

    for (let shtName in current) {
        let c = current[shtName]
        let l = last[shtName]
        let rows = l ? c.filter(r=>!hasSameRow(r, l)) : c.slice()
        for (let [sendto, items] of mergeRows(rows, '抬头')) {
            let cls = eval(shtName)
            let data = items.length == 1 ? items[0] : cls.mergeData(items)
            new cls(data, shtName).exec()
        }
    }

    var workbook = xlsx.readFile('./data.xlsx', {
      type: 'binary'
    });
    xlsx.writeFile(workbook, './data.last.xlsx')

    inputExit('按任意键结束')
} catch (error) {
    inputExit('发生错误：'+error)
}
