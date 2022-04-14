using CentralAptitudeTest.Models;
using System.Windows.Controls;
using System.Windows;
using Microsoft.Win32;
using System.Threading;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ProgressView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProgressView : UserControl
    {
        private Config Config;

        public ProgressView()
        {
            // 대학 엑셀 파일 입력창 시작 시
            InitializeComponent();
            Config = Config.GetConfig();
            ((MainWindow)System.Windows.Application.Current.MainWindow).NextPageButton.IsEnabled = false;
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            // 대학 엑셀 파일 경로 얻기
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == true && openFileDialog.FileName != null)
            {
                Config conf = new Config();
                conf.FilePath = new FilePath() { whole_data_filePath = Config.FilePath.whole_data_filePath, process_data_filePath = openFileDialog.FileName };
                Config.SetConfig(conf);
            }

            // 파일 경로 업로드 딜레이
            Thread.Sleep(500);

            ReadFilePath(openFileDialog.FileName);

            if (!string.IsNullOrEmpty(myTextBox.Text))
            {
                ((MainWindow)System.Windows.Application.Current.MainWindow).NextPageButton.IsEnabled = true;
            }
        }

        private void ReadFilePath(string filename)
        {
            myTextBox.Text = filename;
            return;
        }
    }
}

