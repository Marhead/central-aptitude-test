using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;   // 사용한 엑셀 객체들을 해제 해주기 위한 참조
using CentralAptitudeTest.Models;
using Microsoft.Office.Interop.Excel;   // 액셀 사용을 위한 참조

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        private Config Config;
        private List<Dictionary<string, List<string>>> TempCollegeDictionaries;

        private Application application;

        private Workbook InputWorkbook;
        private Workbook OutputAllWorkbook;
        private Workbook OutputGraphWorkbook;

        private Worksheet InputWorksheet;
        private Worksheet OutputAllWorksheet;
        private Worksheet OutputGraphWorksheet;

        string Datetime = DateTime.Now.ToString("hhmmss");

        public ExcelManipulation(Config config)
        {
            Config = config;
            application = new Application();

            OpenFile(Config.FilePath.whole_data_filePath);

            // Save -> Close 순으로 수행
            SaveFile();
            CloseFile();
        }


        public void OpenFile(string filepath)
        {
            // 입력 Excel 파일(워크북) 불러오기
            InputWorkbook = application.Workbooks.Open(filepath);

            // Excel 화면 창 띄우기
            application.Visible = true;

            // 기존 Excel 파일(워크북) 불러오기
            OutputAllWorkbook = application.Workbooks.Add();
            OutputGraphWorkbook = application.Workbooks.Add();

            // worksheet 생성하기
            InputWorksheet = (Worksheet)InputWorkbook.Sheets[1];
            OutputAllWorksheet = (Worksheet)OutputAllWorkbook.Sheets.Add();
            OutputGraphWorksheet = (Worksheet)OutputGraphWorkbook.Sheets.Add();

        }

        public void SaveFile()
        {
            // Save -> SaveAs 순으로 수행
            InputWorkbook.Save();
            OutputAllWorkbook.Save();
            OutputGraphWorkbook.Save();
            OutputAllWorkbook.SaveAs(Filename: @"C:\\test\\{0}testforall.xlsx", Datetime);
            OutputGraphWorkbook.SaveAs(Filename: @"C:\\test\\{0}testforgraph.xlsx", Datetime);
        }

        public void CloseFile()
        {
            InputWorkbook.Close();
            OutputAllWorkbook.Close();
            OutputGraphWorkbook.Close();
        }

        public string ReadCell(int i, int j)
        {
            Range ColleageName = InputWorksheet.Range["A"];
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
