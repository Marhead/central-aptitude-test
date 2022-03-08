using CentralAptitudeTest.Models;
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;   // 사용한 엑셀 객체들을 해제 해주기 위한 참조
using Microsoft.Office.Interop.Excel;   // 액셀 사용을 위한 참조

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        private Config Config;

        private Application application;

        private Workbook InputWorkbook;
        private Workbook OutputAllWorkbook;
        private Workbook OutputGraphWorkbook;

        private Worksheet InputWorksheet;
        private Worksheet OutputAllWorksheet;
        private Worksheet OutputGraphWorksheet;

        string DesktopPath;
        string Datetime = DateTime.Now.ToString("hhmmss");

        public ExcelManipulation(Config config)
        {
            // 바탕화면 경로 불러오기
            DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Excel 파일 저장 경로 및 파일 이름 설정
            // path = Path.Combine(DesktopPath, "{0}.xlsx", Datetime);

            Config = config;
            application = new Application();

            OpenFile(Config.FilePath.whole_data_filePath);
        }


        public void OpenFile(string filepath)
        {
            // 입력 Excel 파일(워크북) 불러오기
            InputWorkbook = application.Workbooks.Open(filepath);

            Console.WriteLine(InputWorkbook.Worksheets.Count);

            // Excel 화면 창 띄우기
            // application.Visible = true;

            // 기존 Excel 파일(워크북) 불러오기
            OutputAllWorkbook = application.Workbooks.Add();
            OutputGraphWorkbook = application.Workbooks.Add();

            // worksheet 생성하기
            InputWorksheet = (Worksheet)InputWorkbook.Sheets[1];
            OutputAllWorksheet = (Worksheet)OutputAllWorkbook.ActiveSheet;
            OutputGraphWorksheet = (Worksheet)OutputGraphWorkbook.ActiveSheet;

        }

        public void CloseFile()
        {
            // Save -> Close 순으로 수행
            string allfilenaming = "전체" + Datetime + ".xlsx";
            string graphfilenaming = "그래프" + Datetime + ".xlsx";

            string allpath = Path.Combine(DesktopPath, allfilenaming);
            string graphpath = Path.Combine(DesktopPath, graphfilenaming);

            // Save -> SaveAs 순으로 수행
            // InputWorkbook.Save();
            OutputAllWorkbook.SaveAs(Filename: allpath);
            OutputGraphWorkbook.SaveAs(Filename: graphpath);

            InputWorkbook.Close();
            OutputAllWorkbook.Close();
            OutputGraphWorkbook.Close();

            application.Quit();

            // background에서 실행중인 객체들 마저 확실하게 해제시켜주기 위하여 사용.
            Marshal.ReleaseComObject(InputWorkbook);
            Marshal.ReleaseComObject(OutputAllWorkbook);
            Marshal.ReleaseComObject(OutputGraphWorkbook);

            Marshal.ReleaseComObject(application);
        }

        public string ReadCell()
        {
            Range ColleageName = InputWorksheet.UsedRange;

            for(int row = )
            {
                for()
                {

                }
            }

            return "";
        }

        public void WriteToCell()
        {
            Range rg1 = (Range)OutputAllWorksheet.Cells[1, 1];
            rg1.Value = "hello world";
        }
    }
}
