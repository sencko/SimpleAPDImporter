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
            Console.Out.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string filename = "C:\\Users\\i028512\\Documents\\Visual Studio 2017\\Projects\\SimpleAPDImporter\\R83410526971.txt";// "C:\\Users\\i028512\\Documents\\Visual Studio 2017\\Projects\\SimpleAPDImporter\\r83410048802.txt";
            string from = null;
            string to = null;
            if (args.Length > 0)
            {
                filename = args[0];
            }
            if (args.Length > 1)
            {
                from = args[1];
            }
            if (args.Length > 2)
            {
                to = args[2];
            }
            program.readAPD(filename, from, to);
       //   Console.In.ReadLine();
        }
        Regex parts = new Regex(@"^\s\d+(\.\w+)?\s{2,}((\S+\s)*)\s{2,}(((\-?\d*\.\d+)\s*)|(\*{2,}\s*))+$");
        Regex splitter = new Regex(@"\s{2,}");
        Regex newChapter = new Regex(@"^\s{2,}((?:\d+)(?:\.\w+)?)\s((?:\S+\s)*(?:\S)+)$");
        Regex title = new Regex(@"^\s+YTD.+$");
        void readAPD(string filename = "C:\\Users\\i028512\\Documents\\Visual Studio 2017\\Projects\\SimpleAPDImporter\\r83410048802.txt" , string from = null, string to = null)
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
                int fromIndex = 0, toIndex = 0;
                while ((line = reader.ReadLine()) != null)
                {

                    if (title.IsMatch(line))
                    {
                        titleString = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        toIndex = titleString.Length - 1;
                        if (from != null)
                        {
                            fromIndex = Array.IndexOf(titleString, from);
                        }
                        if (to != null)
                        {
                            toIndex = Array.IndexOf(titleString, to);
                        }
                    } else
                    if (newChapter.IsMatch(line))
                    {
                        StartNewSheet(line, newWorkbook, ref sheet, ref i, ref chapterName, ref escapedChapterName, titleString, fromIndex, toIndex);
                    }
                    else
                    {
                        if ((chapterName != null) && (line.StartsWith(chapterName)))
                        {
                            sheet.Cells[2][i].Value2 = "Total Calculated";
                            Excel.Range range = sheet.Range[sheet.Cells[3][i], sheet.Cells[3 + titleString.Length - 1][i]];
                            range.NumberFormat = "0.00";
                            sheet.Cells[3][i].FormulaR1C1 = "=SUM(R3C:R[-1]C)";
                            range.FillRight();
                            i++;
                            sheet.Cells[2][i].Value2 = "Total";
                            SetValues(sheet, i, line.Substring(chapterName.Length), fromIndex, toIndex);

                        }
                        else
                        {

                            Match match = parts.Match(line);
                            if (match.Success)
                            {
                                HandleLine(line, sheet, ref i, fromIndex, toIndex);
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
        
        private void HandleLine( string line, Excel.Worksheet sheet, ref int i, int fromIndex, int toIndex)
        {
            string[] header = splitter.Split(line, 3);

            sheet.Cells[1][i].Value2 = header[0].Trim();
            sheet.Cells[2][i].Value2 = header[1].Trim();

            SetValues(sheet, i, header[2], fromIndex, toIndex);
            i++;
        }

        private static void SetValues(Excel.Worksheet sheet, int i, string intValues, int fromIndex, int toIndex)
        {
            double[] values = intValues.Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Select(x => ((x.Contains('*')) ? (0) : Double.Parse(x))).ToArray();
            Excel.Range range = sheet.Range[sheet.Cells[3][i], sheet.Cells[3 + values.Length - 1][i]];
            range.NumberFormat = "0.00";
            range.Value = values;
        }

        private void StartNewSheet(string line, Excel.Workbook newWorkbook, ref Excel.Worksheet sheet, ref int i, ref string chapterName, ref string escapedChapterName, string[] titleString, int fromIndex, int toIndex)
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
                sheet.Cells[1, 1].Value2 = chapterName.Trim();
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
