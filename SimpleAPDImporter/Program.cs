using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace SimpleAPDImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.readAPD();
            Console.In.ReadLine();
        }

        void readAPD(string filename = "C:\\Users\\i028512\\Documents\\Visual Studio 2017\\Projects\\SimpleAPDImporter\\r83410048802.txt")
        {
            //^\s\d+(\.\w+)?\s{2,}((\S+\s)*)\s{2,}(((\-?\d*\.\d+)\s*)|(\*{2,}\s*))+$
            Regex parts = new Regex(@"^\s\d+(\.\w+)?\s{2,}((\S+\s)*)\s{2,}(((\-?\d*\.\d+)\s*)|(\*{2,}\s*))+");
            Regex splitter = new Regex(@"\s{2,}");

            StreamReader reader = new FileInfo(filename).OpenText();
            string line;
            var ExcelApp = new Excel.Application();
            ExcelApp.
            while ((line = reader.ReadLine()) != null)
            {
                Match match = parts.Match(line);
                if (match.Success)
                {
                  //  int number = int.Parse(match.Groups[1].Value);
                   // string path = match.Groups[2].Value;
                    Console.Out.WriteLine(line);
                    string[] cells = splitter.Split(line);
                    // At this point, `number` and `path` contain the values we want
                    // for the current line. We can then store those values or print them,
                    // or anything else we like.
                }
            }
        }
    }
}
