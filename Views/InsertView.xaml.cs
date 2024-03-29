﻿using CentralAptitudeTest.Models;
using Microsoft.Win32;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using CentralAptitudeTest;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// InsertView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class InsertView : UserControl
    {
        private Config Config;

        public InsertView()
        {
            // 전체 엑셀 파일 입력창 시작 시
            InitializeComponent();
            Config = Config.GetConfig();
            ((MainWindow)Application.Current.MainWindow).NextPageButton.IsEnabled = false;
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            // 전체 엑셀 파일 경로 입력 버튼 클릭 시
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == true && openFileDialog.FileName != null)
            {
                Config conf = new Config();
                conf.FilePath = new FilePath() { whole_data_filePath = openFileDialog.FileName };
                Config.SetConfig(conf);
            }

            // 파일 경로 업로드 시 딜레이
            Thread.Sleep(500);

            ReadFilePath(openFileDialog.FileName);

            if (!string.IsNullOrEmpty(myTextBox.Text))
            {
                ((MainWindow)Application.Current.MainWindow).NextPageButton.IsEnabled = true;
            }
        }

        private void ReadFilePath(string filename)
        {
            myTextBox.Text = filename;
            return;
        }
    }
}
