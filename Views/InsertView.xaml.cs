using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
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
using CentralAptitudeTest.ViewModels;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// InsertView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class InsertView : UserControl
    {
        MainWindow mainWindow = new MainWindow();

        public InsertView()
        {
            InitializeComponent();
            DataContext = new InsertViewModel();
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                //myTextBox.Text = File.ReadAllText(openFileDialog.FileName);
                myTextBox.Text = openFileDialog.FileName;
            }

            //FileStream 클래스 선언하여 진행률 바에 보낼 값 작성
        }

        private void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            root.Content = new ProgressView(myTextBox.Text);
        }
    }
}
