using CentralAptitudeTest.Models;
using Microsoft.Win32;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

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
            InitializeComponent();
            Config = Config.GetConfig();
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == true && openFileDialog.FileName != null)
            {
                Config conf = new Config();
                conf.FilePaths.Add(new FilePath() { filePath = openFileDialog.FileName });
                Config.SetConfig(conf);
            }

            Thread.Sleep(500);

            ReadFilePath(openFileDialog.FileName);
        }

        private void ReadFilePath(string filename) 
        {
            myTextBox.Text = filename;
            return; 
        }
    }
}
