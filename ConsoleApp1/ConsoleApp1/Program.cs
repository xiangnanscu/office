using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
//using System.CodeDom.Compiler;
//using System.CodeDom;

namespace ConsoleApp1
{
    public delegate void DelegateDocument(Word.Document doc);
    public delegate void DelegateWorkbook(Excel.Workbook book);


    class Program
    {
        static string cd = Directory.GetCurrentDirectory();
        static string qfd = cd + "/签发单.doc";
        static Dictionary<string, string> holder;

        static Dictionary<string, string> MakeHolder() {
            return new Dictionary<string, string> {
                {"代字","" },
                {"年份","" },
                {"序号","" },
                {"标题","" },
                {"主送","" },
                {"成文日期","" },
                {"抄送","" },
            };
        }
        static void ProcessWord(string path, DelegateDocument handler)
        {
            Word.Application word = new Word.Application();
            try {             
                var doc = word.Documents.Open(path);
                handler(doc);
            } finally {
                word.Quit();
            }
            
        }
        static void ProcessQfd(string name)
        {
            Word.Application word = new Word.Application();
            try
            {
                var doc = word.Documents.Open(qfd);
                foreach (KeyValuePair<string, string> kvp in holder) {
                    Word.Find findObject = word.Application.Selection.Find;
                    findObject.Text = "{"+kvp.Key+"}";
                    findObject.Replacement.Text = kvp.Value;
                    findObject.Execute(Replace: Word.WdReplace.wdReplaceAll);
                }
                doc.SaveAs(FileName: cd + $"/{name}.doc");
            }
            finally
            {
                word.Quit();
            }

        }
        static void ProcessExcel(string path, DelegateWorkbook handler)
        {
            Excel.Application excel = new Excel.Application();
            try
            {               
                var xls = excel.Workbooks.Open(path);
                handler(xls);
            }
            finally
            {
                excel.Quit();
            }

        }



        private static string ToLiteral(string input)
        {
            //using (var writer = new StringWriter())
            //{
            //    using (var provider = CodeDomProvider.CreateProvider("CSharp"))
            //    {
            //        provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
            //        return writer.ToString();
            //    }
            //}
            return input;
        }

        static void Print(params string[] strs)
        {
            foreach (var item in strs)
            {
                Console.WriteLine(ToLiteral(item));
            }
        }
        static void GetWenHao(Word.Document document) {
            var pattern = @"(?<代字>\w+)〔(?<年份>\d{4})〕(?<序号>\d+)号";
            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, pattern);
                if (m.Success)
                {
                    holder["代字"] = m.Groups["代字"].Value;
                    holder["年份"] = m.Groups["年份"].Value;
                    holder["序号"] = m.Groups["序号"].Value;
                    return;
                }
            }
            Print("警告:没有找到文号");
        }
        static void GetBiaoTi(Word.Document document)
        {
            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                if (r.Font.Name == "方正小标宋简体" && r.Font.Size == 22)
                {
                    holder["标题"] = holder["标题"] + Regex.Replace(r.Text, "[\n\r ]", "");
                }
            }
        }
        static void GetZhuSong(Word.Document document)
        {
            var pattern = @"^\s*([^:：\s]+)[:：]\s*$";
            
            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, pattern);
                if (m.Success &&
                    r.Font.Name == "仿宋_GB2312" && r.Font.Size == 16 
                    && p.FirstLineIndent == 0
                  )
                {
                    holder["主送"] = m.Groups[1].Value;
                    return;
                }
            }
            Print("警告:没有找到主送机关");
        }
        static void GetDate(Word.Document document)
        {
            var pattern = @"^\s*(\d+年\d+月\d+日)\s*$";

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, pattern);
                if (m.Success && r.Font.Name == "仿宋_GB2312" && r.Font.Size == 16)
                {
                    holder["成文日期"] = m.Groups[1].Value;
                }
            }
        }
        static void GetChaoSong(Word.Document document)
        {
            var pattern = @"^\s*抄送[:：]([^:：\s]+)\s*$";

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, pattern);
                if (m.Success && r.Font.Name == "仿宋_GB2312" && r.Font.Size < 16)
                {
                    holder["抄送"] = m.Groups[1].Value;
                    return;
                }
            }
        }
        static void HandleFile(Word.Document document)
        {
            GetWenHao(document);
            GetBiaoTi(document);
            GetZhuSong(document);
            GetDate(document);
            GetChaoSong(document);
            document.Close(false);
        }



        static void Work() {
            var pat = new Regex(@"^[^~]+\.docx?$");
            var docs = Directory.GetFiles(cd, "*.*", SearchOption.AllDirectories)
                .Where(file => pat.IsMatch(file))
                .ToList();
            foreach (string docPath in docs)
            {
                if (docPath.Contains("签发单"))
                {
                    continue;
                }
                Print("处理文件:" + docPath);
                holder = MakeHolder();
                ProcessWord(docPath, new DelegateDocument(HandleFile));
                ProcessQfd($"签发单-{(holder["代字"])}-{(holder["年份"])}-{(holder["序号"])}号");

            }

        }
        static void Test()
        {

        }
        static void Main(string[] args)
        {
            try
            {
                Print("------高大上的签发单生成程序开始^_^------");
                Work();
            }
            finally
            {
                
                Print("------程序结束, 按任意键退出------");
                Console.Read();
            }
  

        }
    }
}
