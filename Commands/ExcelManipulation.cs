using CentralAptitudeTest.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;   // 사용한 엑셀 객체들을 해제 해주기 위한 참조
using Microsoft.Office.Interop.Excel;   // 액셀 사용을 위한 참조

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        private Config Config;
        private List<Dictionary<string, List<string>>> TempCollegeDictionaries;

        private Application InputApplication;
        private Application OutputAllApplication;
        private Application OutputGraphApplication;

        private Workbook InputWorkbook;
        private Workbook OutputAllWorkbook;
        private Workbook OutputGraphWorkbook;

        private Worksheet InputWorksheet;
        private Worksheet OutputAllWorksheet;
        private Worksheet OutputGraphWorksheet;

        string Datetime = DateTime.Now.ToString("hhmmss");

        public ExcelManipulation(Config config)
        {
            // 바탕화면 경로 불러오기
            string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Excel 파일 저장 경로 및 파일 이름 설정
            string path = Path.Combine(DesktopPath, "전체.xlsx");

            Config = config;
            InputApplication = new Application();
            OutputAllApplication = new Application();
            OutputGraphApplication = new Application();

            OpenFile(Config.FilePath.whole_data_filePath);

            // Save -> Close 순으로 수행
            SaveFile();
            CloseFile();
        }


        public void OpenFile(string filepath)
        {
            // 입력 Excel 파일(워크북) 불러오기
            InputWorkbook = InputApplication.Workbooks.Open(filepath);

            Console.WriteLine(InputWorkbook.Worksheets.Count);

            // Excel 화면 창 띄우기
            InputApplication.Visible = true;

            // 기존 Excel 파일(워크북) 불러오기
            OutputAllWorkbook = OutputAllApplication.Workbooks.Add();
            OutputGraphWorkbook = OutputGraphApplication.Workbooks.Add();

            // worksheet 생성하기
            InputWorksheet = (Worksheet)InputWorkbook.Sheets[1];
            OutputAllWorksheet = (Worksheet)OutputAllWorkbook.Sheets.Add();
            OutputGraphWorksheet = (Worksheet)OutputGraphWorkbook.Sheets.Add();

        }

        public void SaveFile()
        {
            // Save -> SaveAs 순으로 수행
            InputWorkbook.Save();
            OutputAllWorkbook.SaveAs(Filename: "C:\\test\\{0} testforall.xlsx", Datetime);
            OutputGraphWorkbook.SaveAs(Filename: "C:\\test\\{0} testforgraph.xlsx", Datetime);
        }

        public void CloseFile()
        {
            InputWorkbook.Close();
            OutputAllWorkbook.Close();
            OutputGraphWorkbook.Close();

            InputApplication.Quit();
            OutputAllApplication.Quit();
            OutputGraphApplication.Quit();

            // background에서 실행중인 객체들 마저 확실하게 해제시켜주기 위하여 사용.
            Marshal.ReleaseComObject(InputWorkbook);
            Marshal.ReleaseComObject(OutputAllWorkbook);
            Marshal.ReleaseComObject(OutputGraphWorkbook);

            Marshal.ReleaseComObject(InputApplication);
            Marshal.ReleaseComObject(OutputAllApplication);
            Marshal.ReleaseComObject(OutputGraphApplication);

        }

        public string ReadCell()
        {
            Range ColleageName = InputWorksheet.Range["A"];

            return "";
        }

        public void WriteToCell()
        {

        }

        private void SortingCells()
        {
            // workbook에 데이터를 읽거나 쓸 worksheet 생성
        }

        private void Filtering1()
        {

        }

        private void Filtering2()
        {

        }
        private void Filtering3()
        {

        }
        private void Filtering4()
        {

        }
        private void Filtering5()
        {

        }
        private void Filtering6()
        {

        }
        private void Filtering7()
        {

        }
        private void Filtering8()
        {

        }
        private void Filtering9()
        {

        }
        private void Filtering10()
        {

        }
        private void Filtering11()
        {

        }

        private void Filtering12()
        {

        }
        private void Filtering13()
        {

        }
        private void Filtering14()
        {

        }
        private void Filtering15()
        {

        }
        private void Filtering16()
        {

        }
    }
}
