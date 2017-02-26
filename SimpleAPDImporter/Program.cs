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
        Regex parts = new Regex(@"^\s\d+(\.\w+)?\s{2,}((\S+\s)*)\s{2,}(((\-?\d*\.\d+)\s*)|(\*{2,}\s*))+");
        Regex splitter = new Regex(@"\s{2,}");
        Regex newChapter = new Regex(@"^\s{2,}((?:\d+)(?:\.\w+)?)\s((?:\S+\s)*(?:\S)+)$");
        Regex title = new Regex(@"^\s+YTD.+$");
        void readAPD(string filename = "C:\\Users\\i028512\\Documents\\Visual Studio 2017\\Projects\\SimpleAPDImporter\\r83410048802.txt")
        {
            //^\s\d+(\.\w+)?\s{2,}((\S+\s)*)\s{2,}(((\-?\d*\.\d+)\s*)|(\*{2,}\s*))+$


            StreamReader reader = new FileInfo(filename).OpenText();
            string line;
            Excel.Application application = new Excel.Application();
            try
            {
              
                application.ScreenUpdating = false;

                Excel.Workbook newWorkbook = application.Workbooks.Add();
                Excel.Worksheet sheet = null;
                application.Calculation = Excel.XlCalculation.xlCalculationManual;
                int i = 3;
                string chapterName = null;
                string escapedChapterName = null;
                string[] titleString = null;
                while ((line = reader.ReadLine()) != null)
                {

                    if (title.IsMatch(line))
                    {
                        titleString = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    } else
                    if (newChapter.IsMatch(line))
                    {
                        StartNewSheet(line, newWorkbook, ref sheet, ref i, ref chapterName, ref escapedChapterName, titleString);
                    }
                    else
                    {
                        if ((chapterName != null) && (line.StartsWith(chapterName)))
                        {
                            // total check value

                        }
                        else
                        {

                            Match match = parts.Match(line);
                            if (match.Success)
                            {
                                HandleLine(line, sheet, ref i);
                            }
                        }
                    }
                }
                SetCellSize(sheet);
            }
            finally
            {
                application.Visible = true;
                application.ScreenUpdating = true;
                application.Calculate();
                application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
        }
        
        private void HandleLine( string line, Excel.Worksheet sheet, ref int i)
        {
            string[] header = splitter.Split(line, 3);

            sheet.Cells[1][i].Value2 = header[0].Trim();
            sheet.Cells[2][i].Value2 = header[1].Trim();

            double[] values = header[2].Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Select<string, double>(x => ((x.Contains('*')) ? (0) : Double.Parse(x))).ToArray<double>();
            Excel.Range range = sheet.Range[sheet.Cells[3][i], sheet.Cells[3 + values.Length - 1][i]];
            range.NumberFormat = "0.00";
            range.Value = values;
            i++;
        }

        private void StartNewSheet(string line, Excel.Workbook newWorkbook, ref Excel.Worksheet sheet, ref int i, ref string chapterName, ref string escapedChapterName, string[] titleString)
        {
            if ((chapterName == null) || !chapterName.Equals(line))
            {
                SetCellSize(sheet);
                chapterName = line;
                Console.Out.WriteLine(chapterName);
                sheet = newWorkbook.Worksheets.Add();
                escapedChapterName = escape(chapterName);
                sheet.Name = escapedChapterName;
                sheet.Range["A:A"].NumberFormat = "@";
                sheet.Range["B:B"].NumberFormat = "@";
                sheet.Cells[1, 1].Value2 = chapterName;
                sheet.Range[sheet.Cells[1][1], sheet.Cells[2][1]].Merge();
                sheet.Range[sheet.Cells[3][1], sheet.Cells[3 + titleString.Length - 1][1]].Value2 = titleString;

                sheet.Cells[3][2].Value2 = "PLN";
                sheet.Range[sheet.Cells[3][2], sheet.Cells[3 + titleString.Length - 1][2]].FillRight();
                i = 3;
            }
        }

        private static void SetCellSize(Excel.Worksheet sheet)
        {
            if (sheet != null)
            {
                sheet.UsedRange.Columns.AutoFit();
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
