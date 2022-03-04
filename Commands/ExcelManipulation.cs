using System;
using System.Collections.Generic;
using System.Text;
using CentralAptitudeTest.Models;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace CentralAptitudeTest.Commands
{
    class ExcelManipulation
    {
        private Config Config;
        private List<Dictionary<string, List<string>>> TempCollegeDictionaries;

        private Application application;
        
        private Workbook InputWorkbook;
        private Workbook OutputWorkbookForAll;
        private Workbook OutputWorkbookForGraph;

        private Sheets WorksheetInput;
        private Sheets WorksheetOutputAll;
        private Sheets WorksheetOutputGraph;

        // 새로운 Excel 파일(워크북) 생성
        // Workbook workbook = application.Workbooks.Add();

        public ExcelManipulation(Config config)
        {
            Config = config;
            application = new Application();

            OpenFile(Config.FilePath.filePath);
            SaveFile();
            CloseFile();
        }


        public void OpenFile(string filepath)
        {
            // 입력 Excel 파일(워크북) 불러오기
            InputWorkbook = application.Workbooks.Open(Filename: @filepath);
            application.Visible = true;

            // 기존 Excel 파일(워크북) 불러오기
            OutputWorkbookForAll = application.Workbooks.Add();
            OutputWorkbookForGraph = application.Workbooks.Add();

            // worksheet 생성하기
            WorksheetInput = InputWorkbook.Sheets;
            WorksheetOutputAll = OutputWorkbookForAll.Sheets;
            WorksheetOutputGraph = OutputWorkbookForGraph.Sheets;

        }

        public void SaveFile()
        {
            OutputWorkbookForAll.Save();
            OutputWorkbookForGraph.Save();
            OutputWorkbookForAll.SaveAs(Filename : @"C:\\test\\testforall.xlsx");
            OutputWorkbookForGraph.SaveAs(Filename : @"C:\\test\\testforgraph.xlsx");
        }

        public void CloseFile()
        {
            InputWorkbook.Close();
            OutputWorkbookForAll.Close();
            OutputWorkbookForGraph.Close();
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
