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
            Regex newChapter = new Regex(@"^\s{2,}((?:\d+)(?:\.\w+)?)\s((?:\S+\s)*(?:\S)+)$");

            StreamReader reader = new FileInfo(filename).OpenText();
            string line;
            Excel.Application application = new Excel.Application();
            application.Visible = true;
            Excel.Workbook newWorkbook = application.Workbooks.Add();
            Excel.Worksheet sheet = null;
            int i = 0;
            string chapterName = null;
            string escapedChapterName = null;
            while ((line = reader.ReadLine()) != null)
            {
                if (newChapter.IsMatch(line))
                {
                    if ((chapterName == null) || !chapterName.Equals(line))
                    {
                        chapterName = line;
                        sheet = newWorkbook.Worksheets.Add();
                        escapedChapterName = escape(chapterName);
                        sheet.Name = escapedChapterName;
                        sheet.Activate();

                    }
                }
                else
                {
                    if ((chapterName!=null) && (line.StartsWith(chapterName)))
                    {
                        // total check value

                    }
                    else
                    {

                        Match match = parts.Match(line);
                        if (match.Success)
                        {
                            Console.Out.WriteLine(sheet.Name);
                            string[] header = splitter.Split(line, 3);
                            
                            sheet.Cells[1][i + 1].Value = header[0].Trim();
                            sheet.Cells[2][i + 1].Value = header[1].Trim();

                            string[] values = header[2].Split(new char[]{ ' ', '\t'}, StringSplitOptions.RemoveEmptyEntries);
                            for (int j = 0; j < values.Length; j++)
                            {
                                sheet.Cells[j + 3][i + 1] = values[j].Trim();
                            }
                            i++;
                        }
                    }
                }
            }
        }

        private string escape(string chapterName)
        {
            if (chapterName == null)
            {
                return "null";
            }
            string ret = chapterName.Replace(':', ' ').Replace('\\', ' ').Replace('/', ' ').Replace('?', ' ').Replace('*', ' ').Replace('[', ' ').Replace(']', ' ').Trim();
            if (ret.Length > 30)
            {
                ret = ret.Substring(0, 30);
            }
            return ret;
        }
    }
}
