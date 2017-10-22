using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static Word.Application word = new Word.Application();
        static Excel.Application excel = new Excel.Application();
        public Program()
        {
            Print(@"init...");
        }
        ~Program()
        {
            Print(@"end...");
            word.Quit();
            excel.Quit();
            Print(@"按任意键退出");
            Console.Read();
        }
        static void Print(string text)
        {
            Console.WriteLine(text);
        }
        static void ProcessDoc(string path)
        {
            Print(path);
            var document = word.Documents.Open(path);

            foreach (Word.Paragraph p in document.Paragraphs)
            {
                // Write the word.
                var pattern = @"(?<代字>\w+)〔(?<年份>\d{4})〕(?<序号>\d+)号";
                var r = p.Range;
                Console.WriteLine($"{r.Font.Name},{r.Font.Size},首行缩进:{p.FirstLineIndent}");
                var m = Regex.Match(r.Text, pattern);
                if (m.Success)
                {
                    // Console.WriteLine($"{m.Groups[\"文号\"]},{m.Groups[\"年份\"]},{m.Groups[\"号\"]}");
                    //break;
                }
                // Console.WriteLine(r.Text);
                // Console.WriteLine("-----------------");

            }

 
            document.Save();
        }
        static void Main(string[] args)
        {
            var cd = Directory.GetCurrentDirectory();
            var pat = new Regex(@"^[^~]+\.docx?$");
            var docs = Directory.GetFiles(cd, "*.*", SearchOption.AllDirectories)
                .Where(file => pat.IsMatch(file))
                .ToList();
            foreach (string doc in docs)
            {
                ProcessDoc(doc);
                Print(doc);
            }
            word.Quit();
            excel.Quit();
            Console.Read();

        }
        static void ReadQianFaDan()
        {

        }
    }
}
