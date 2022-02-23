using CentralAptitudeTest.Models;
using System.Collections.Generic;
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
        private List<Dictionary<string, List<string>>> Temp_College_Dictionarys;

        public ProgressView()
        {
            InitializeComponent();
            Config = Config.GetConfig();
        }

        private void UploadButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Config conf = new Config();

            conf.FilePath = new FilePath()
            {
                filePath = Config.FilePath.filePath,
                College_Dictionarys = Temp_College_Dictionarys,
            };
            Config.SetConfig(conf);
            return;
        }

        private void AddCollegeButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 단과대 정보 추가
            if (Config.FilePath != null)
            {
                Config conf = new Config();

                conf.Subjects = new List<string> { subject1.Text, subject2.Text, subject3.Text, subject4.Text, subject5.Text, subject6.Text };
                Dictionary<string, List<string>> dictionary = new Dictionary<string, List<string>>() {
                        { college.Text, conf.Subjects },
                    };

                if (Temp_College_Dictionarys != null)
                {
                    Temp_College_Dictionarys.Add(dictionary);
                }
                else
                {
                    Temp_College_Dictionarys = new List<Dictionary<string, List<string>>>() { { dictionary }, };
                }

                foreach (string key in dictionary.Keys)
                {
                    this.college_combo.Items.Add(key);
                }
                return;
            }
        }

        private void Input_Complete_Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 치훈이 함수 호출 
        }

        //
        //private void OpenFile(string filepath)
        //{
        //    var ExcelApp = new Excel.Application();
        //    ExcelApp.Visible = true;

        //    Excel.Workbooks books = ExcelApp.Workbooks;

        //    // FilePath 넣는 곳
        //    Excel.Workbook sheets = books.Open(filepath);
        //}

        //private void illustrates()
        //{
        //    Application excelApplication = new Application();

        //    Excel.Workbook excelWorkBook = Excel.Workbooks.Open(@"D:\Test.xslx");

        //    int worksheetcount = excelWorkBook.Worksheets.Count;
        //    if (worksheetcount > 0)
        //    {
        //        Excel.Worksheet worksheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
        //        string worksheetName = worksheet.Name;
        //        var data = ((Excel.Range)worksheet.Cells[1050, 20]).Value;
        //        Console.WriteLine(data);
        //    }
        //    else
        //    {
        //        Console.WriteLine("No worksheets available");
        //    }
        //}
    }
}
