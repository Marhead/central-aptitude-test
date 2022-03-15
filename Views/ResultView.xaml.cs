using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;
using CentralAptitudeTest.Commands;
using CentralAptitudeTest.Models;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ResultView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ResultView : UserControl
    {
        private Config config;

        public ResultView()
        {
            InitializeComponent();
            config = Config.GetConfig();
        }

        private void AddCollegeButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("엑셀 생성자 생성...");
            //ExcelManipulation 함수 호출
            ExcelManipulation ExcelManipulation = new ExcelManipulation(config);

            Console.WriteLine("단과대, 학과 읽기 시작...");
            ExcelManipulation.ReadCollege();

            // ExcelManipulation.WriteToCell();

            // MessageBox.Show(ExcelManipulation.ReadCollege());

            Console.WriteLine("작업 완료, 파일 닫기 시작...");
            ExcelManipulation.CloseFile();
            MessageBox.Show("변환 완료!");
        }
    }
}
