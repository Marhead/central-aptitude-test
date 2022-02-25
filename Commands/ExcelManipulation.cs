using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        private void OpenFile(string filepath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = true;

            Excel.Workbooks books = ExcelApp.Workbooks;

            // FilePath 넣는 곳
            Excel.Workbook sheets = books.Open(filepath);
        }

        private void illustrates()
        {
            Excel.Application excelApplication = new Excel.Application();

            Excel.Workbook excelWorkBook = Excel.Workbooks.Open(@"D:\Test.xslx");

            int worksheetcount = excelWorkBook.Worksheets.Count;
            if (worksheetcount > 0)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
                string worksheetName = worksheet.Name;
                var data = ((Excel.Range)worksheet.Cells[1050, 20]).Value;
                Console.WriteLine(data);
            }
            else
            {
                Console.WriteLine("No worksheets available");
            }
        }
    }
}
