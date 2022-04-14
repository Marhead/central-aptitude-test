using System;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
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
        private ExcelManipulation excelManipulation;

        public ResultView()
        {
            InitializeComponent();
            config = Config.GetConfig();
        }

        private void AddCollegeButton_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.WorkerReportsProgress = true;

            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.RunWorkerAsync();            
        }

        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
            ProgressTextBlock.Text = (string)e.UserState;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, String.Format("엑셀 조작 시작"));

            //ExcelManipulation 함수 호출
            ExcelManipulation ExcelManipulation = new ExcelManipulation(config, worker);
            excelManipulation = ExcelManipulation;

            excelManipulation.ReadCollege();

            try
            {
                excelManipulation.MisfitFiltering();
            }
            catch (Exception exception)
            {
                MessageBox.Show("필터링 중 오류발생!!!\n다시 작동 시켜주세요!");
            }

            try
            {
                excelManipulation.SeparateEachDepart();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Excel 입력 데이터 오류 발견!!!\n데이터를 수정하고 다시 작동 시켜주세요!");
            }

            try
            {
                excelManipulation.GraphFileTask();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Excel graph 결과 출력 오류!!!\n데이터를 수정하고 다시 작동 시켜주세요!");
            }

            try
            {
                excelManipulation.ResultEachCollege();
            }
            catch (Exception exception)
            {
                MessageBox.Show("각 엑셀시트 결과 산출 오류!!!\n데이터를 수정하고 다시 작동 시켜주세요!");
            }


            excelManipulation.CloseFile();
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("변환 완료!");
            ProgressBar.Value = 0;
            ProgressTextBlock.Text = "변환 완료!!!";
        }
    }
}