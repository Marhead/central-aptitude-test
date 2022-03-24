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
            //ExcelManipulation 함수 호출
            ExcelManipulation ExcelManipulation = new ExcelManipulation(config);
            ExcelManipulation.ReadCollege();

            try
            {
                ExcelManipulation.SeparateEachDepart();
            }catch(Exception exception)
            {
                MessageBox.Show("Excel 입력 데이터 오류 발견!!!\n데이터를 수정하고 다시 작동 시켜주세요!");
            }

            try
            {
                ExcelManipulation.GraphFileTask();
            }catch(Exception exception)
            {
                MessageBox.Show("Excel graph 결과 출력 오류!!!\n데이터를 수정하고 다시 작동 시켜주세요!");
            }

            ExcelManipulation.CloseFile();
            MessageBox.Show("변환 완료!");
        }
    }
}
