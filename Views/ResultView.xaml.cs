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

            // 1번째 작업
            try
            {
                excelManipulation.ReadCollege();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + "\n학과 데이터에 오류가 있습니다. 학과데이터 파일을 확인해주세요.");
            }

            // 2번째 작업
            // excelManipulation.MisfitFiltering();
            try
            {
                excelManipulation.MisfitFiltering();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }

            // 3번째 작업
            // excelManipulation.SeparateEachDepart();
            try
            {
                var ret = excelManipulation.SeparateEachDepart();
                if(ret != null)
                {
                    MessageBox.Show("일람표 데이터 중, 학과표에 없는 데이터가 있습니다.\n학과 데이터 확인 해주세요.\n문제 학과 : " + ret);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + "\n학과 데이터와 일람표 데이터 간에 차이가 있습니다. 확인 부탁드립니다.");
            }

            // 4번째 작업
            // excelManipulation.GraphFileTask();
            try
            {
                excelManipulation.GraphFileTask();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + "\n그래프 파일 작성 중 문제가 발생했습니다.\n일람표 워크시트 명을 확인해 주세요.");
            }

            // 5번째 작업
            // excelManipulation.ResultEachCollege();
            try
            {
                excelManipulation.ResultEachCollege();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + "\n결과 데이터 작성 중 문제가 발생하였습니다.\n일람표 워크시트 명을 확인해 주세요.");
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