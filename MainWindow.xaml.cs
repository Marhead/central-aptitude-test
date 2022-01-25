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
using CentralAptitudeTest.Views;

namespace CentralAptitudeTest
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void panelHeader_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void StackPanel1_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //asd.Content = new InsertView();
        }

        private void StackPanel2_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {            
            //asd.Content = new ProgressView(null);
        }

        private void StackPanel3_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //asd.Content = new ResultView();
        }

        private void Insert_Selected(object sender, RoutedEventArgs e)
        {
            asd.Content = new InsertView();
        }

        private void Process_Selected(object sender, RoutedEventArgs e)
        {

        }

        private void ListViewItem_Selected(object sender, RoutedEventArgs e)
        {

        }
    }
}
