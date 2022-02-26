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
        private Config config;
        private Application application;
        
        private Workbook InputWorkbook = new Workbook();
        private Worksheet Worksheet;

        private Workbook OutputWorkbookForAll = new Workbook();
        private Workbook OutputWorkbookForGraph = new Workbook();

        // 새로운 Excel 파일(워크북) 생성
        // Workbook workbook = application.Workbooks.Add();

        public ExcelManipulation(string filepath)
        {
            application = new Application();
            OpenFile(filepath);
        }


        public void OpenFile(string filepath)
        {
            // Config에서 filepath 뽑아오기
            //string filepath = config.FilePath.filePath;

            // 기존 Excel 파일(워크북) 불러오기
            InputWorkbook = application.Workbooks.Open(Filename : filepath);
            OutputWorkbookForAll = application.Workbooks.Add();
            OutputWorkbookForGraph = application.Workbooks.Add();
        }

        public void SaveFile()
        {
            OutputWorkbookForAll.SaveAs(Filename : @"C:\\test\testforall.xlsx");
            OutputWorkbookForGraph.SaveAs(Filename : @"C:\\test\testforgraph.xlsx");
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
