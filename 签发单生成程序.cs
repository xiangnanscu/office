using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Marshal = System.Runtime.InteropServices.Marshal;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace ConsoleApp1
{

    class Program
    {
        static Word.Application word;
        static Excel.Application excel;
        static Dictionary<string, string> holder;
        static Boolean isDoc = false;
        static string cd = Directory.GetCurrentDirectory();
        static string qianfadan;
        static HashSet<string> fileTypes = new HashSet<string>() { ".doc", ".docx", ".xls", ".xlsx" };
        static string patternChaosong = @"^\s*抄送[:：]([^:：\n\r。.]+)[。.]?\s*";
        static string patternChun = @"\s*存[:：]([^:：\n\r。.]+)[。.]?";
        static string patternZhuSong = @"^\s*([^:：\s]+)[:：]\s*$";
        static string patternWenHao = @"^\s*(?<代字>\w+)〔(?<年份>\d{4})〕(?<序号>[\d ]*)号\s*";
        static string patternDate = @"^\s*(\d+年\d+月\d+日)\s*$";
        static string qianfadanPrefix = "签发单-";

        static Dictionary<string, string> MakeHolder()
        {
            return new Dictionary<string, string> {
                {"代字","" },
                {"年份","" },
                {"序号","" },
                {"标题","" },
                {"主送","" },
                {"成文日期","" },
                {"抄送","" },
                {"存","" },
            };
        }

        private static void DeleteQianFaDan()
        {
            var candidates = Directory.GetFiles(cd, "*.*", SearchOption.AllDirectories)
                .Where(file => Path.GetFileNameWithoutExtension(file).StartsWith(qianfadanPrefix) && fileTypes.Contains(Path.GetExtension(file)))
                .ToList();
            foreach (var file in candidates) {
                File.Delete(file);
            }
            
        }

        private static void FindQianFaDan()
        {
            var candidates = Directory.GetFiles(cd, "*.*", SearchOption.AllDirectories)
                .Where(file => Path.GetFileNameWithoutExtension(file) == "签发单" && fileTypes.Contains(Path.GetExtension(file)))
                .ToList();
            if (candidates.Count == 0)
            {
                throw new Exception("没有找到签发单模板文件");
            }
            else
            {
                qianfadan = candidates[0];
                var t = Path.GetExtension(qianfadan);
                if (t == ".doc" || t == ".docx")
                {
                    isDoc = true;
                }
                Print("找到签发单模板文件:" + Path.GetFileName(qianfadan));
            }
        }


        private static string ToLiteral(string input)
        {
            return input;
        }

        static void Print(params string[] strs)
        {
            foreach (var item in strs)
            {
                Console.WriteLine(ToLiteral(item));
            }
        }
        static void GetWenHao(Word.Document document)
        {

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, patternWenHao);
                if (m.Success)
                {
                    holder["代字"] = m.Groups["代字"].Value;
                    holder["年份"] = m.Groups["年份"].Value;
                    holder["序号"] = m.Groups["序号"].Value;
                    return;
                }
            }
            Print("警告:没有找到文号，请确保格式类似“江人社发〔2017〕1号”");
        }
        static void GetBiaoTi(Word.Document document)
        {
            var find = false;
            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                if (r.Font.Name == "方正小标宋简体" && r.Font.Size == 22)
                {
                    find = true;
                    holder["标题"] = holder["标题"] + Regex.Replace(r.Text, "[\n\r ]|(江安县人力资源和社会保障局)", "");
                }
            }
            if (!find)
            {
                Print("警告：没有找到标题，请确保字体为“方正小标宋简体”，字号为二号");
            }

        }
        static void GetZhuSong(Word.Document document)
        {

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, patternZhuSong);
                if (m.Success && r.Font.Name == "仿宋_GB2312" && r.Font.Size == 16
                    && p.FirstLineIndent == 0)
                {
                    holder["主送"] = m.Groups[1].Value;
                    return;
                }
            }
            Print("警告:没有找到主送机关，请确保字体为“仿宋_GB2312”，字号为三号，无首行缩进");
        }
        static void GetDate(Word.Document document)
        {
            var find = false;
            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, patternDate);
                if (m.Success && r.Font.Name == "仿宋_GB2312" && r.Font.Size == 16)
                {
                    holder["成文日期"] = m.Groups[1].Value;
                    find = true;
                }
            }
            if (!find)
            {
                Print("警告：没有找到成文日期，请确保格式类似于“2017年10月20日”，字体为“仿宋_GB2312”，字号为三号");
            }
        }
        static void GetChaoSong(Word.Document document)
        {

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, patternChaosong);
                if (m.Success && r.Font.Name == "仿宋_GB2312")
                {
                    holder["抄送"] = "抄送：" + m.Groups[1].Value;
                    return;
                }
            }
            Print("警告：没有找到抄送机关，请确保格式类似于“抄送：县委组织部”，字体为“仿宋_GB2312”，字号小于三号");
        }
        static void GetChun(Word.Document document)
        {

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                var r = p.Range;
                var m = Regex.Match(r.Text, patternChun);
                if (m.Success && r.Font.Name == "仿宋_GB2312")
                {
                    holder["存"] = "存：" + m.Groups[1].Value;
                    return;
                }
            }
        }
        static void Work()
        {
            var pat = new Regex(@"^[^~]+\.docx?$");
            var docs = Directory.GetFiles(cd, "*.*", SearchOption.AllDirectories)
                .Where(file => pat.IsMatch(file))
                .ToList();
            if (docs.Count == 0)
            {
                Print("警告：没有找到word文件");
            }
            foreach (string docPath in docs)
            {
                var name = Path.GetFileNameWithoutExtension(docPath);
                var modifyTime = File.GetLastWriteTime(docPath);
                if (DateTime.Now > modifyTime.AddMinutes(5)) {
                    continue;
                }
                Print($"处理文件:{Path.GetFileName(docPath)}");
                holder = MakeHolder();
                word = new Word.Application();
                try
                {
                    var doc = word.Documents.Open(docPath);
                    GetWenHao(doc);
                    GetBiaoTi(doc);
                    GetZhuSong(doc);
                    GetDate(doc);
                    GetChaoSong(doc);
                    GetChun(doc);
                    doc.Close(false);
                }
                finally
                {
                    word.Quit();
                }
                ProcessQianFaDan($"{qianfadanPrefix}{name}");
            }
        }
        static void ProcessQianFaDan(string name)
        {
            if (isDoc)
            {
                word = new Word.Application();
                try
                {

                    var doc = word.Documents.Open(qianfadan);
                    foreach (KeyValuePair<string, string> kvp in holder)
                    {
                        Word.Find findObject = word.Application.Selection.Find;
                        findObject.Text = "{" + kvp.Key + "}";
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
            else
            {
                Excel.Application excel = new Excel.Application();

                try
                {
                    excel.DisplayAlerts = false;
                    Excel.Workbooks wbs = excel.Workbooks;
                    Excel.Workbook book = wbs.Open(qianfadan);

                    var sht = book.ActiveSheet;
                    Excel.Range r = sht.UsedRange;
                    
                    foreach (KeyValuePair<string, string> kvp in holder)
                    {
                        r.Replace("{" + kvp.Key + "}", kvp.Value);
                    }
                    book.SaveAs(cd + $"/{name}.xls");
                    excel.DisplayAlerts = true;
                    book.Close(false);
                    wbs.Close();
                    
                    excel.Quit();
                    Marshal.FinalReleaseComObject(r);
                    Marshal.FinalReleaseComObject(sht);
                    Marshal.FinalReleaseComObject(book);
                    Marshal.FinalReleaseComObject(wbs);
                    Marshal.FinalReleaseComObject(excel);
                }
                catch (Exception e)
                {
                    Print("发生错误："+e.ToString());
                }
                finally
                {

                    

                }
            }
        }
        static void Main(string[] args)
        {
            try
            {
                Print("------高大上的签发单生成程序开始^_^------");
                DeleteQianFaDan();
                FindQianFaDan();
                Work();
            }
            catch (Exception ex)
            {
                Print("出现异常：", ex.ToString());
            }
            finally
            {

                Print("------程序结束, 按任意键退出-------------");
                Console.Read();
            }


        }
    }
}
