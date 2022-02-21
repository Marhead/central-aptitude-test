using CentralAptitudeTest.Models;
using System.Windows.Controls;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ProgressView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProgressView : UserControl
    {
        private Config Config;

        private Excel.Application excelApp = new Excel.Application();

        public ProgressView()
        {
            InitializeComponent();

            Config = Config.GetConfig();

            foreach (FilePath fp in Config.FilePaths)
            {
                if (fp.filePath != null)
                {
                    college.Content = fp.College;
                    subject.Content = fp.Subject;
                    TextBlock.Text = fp.filePath;
                }
            }

            // OpenFile 호출, Config에서 주소 불러와서 대입
            OpenFile(Config.FilePaths);

            return;
        }

        //
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
            Application excelApplication = new Application();

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
