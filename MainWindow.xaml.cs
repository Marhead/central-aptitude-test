using System.Windows;
using System.Windows.Input;
using CentralAptitudeTest.Models;
using CentralAptitudeTest.Views;

namespace CentralAptitudeTest
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private Config Config;
        private int PageNum = 0;

        public MainWindow()
        {
            InitializeComponent();
            MainControl.Content = new InsertView();
            this.Insert.IsSelected = true;    
        }

        private void panelHeader_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Insert_Selected(object sender, RoutedEventArgs e)
        {
            //asd.Content = new InsertView();
        }

        private void Process_Selected(object sender, RoutedEventArgs e)
        {
            //asd.Content = new ProgressView(String.Empty);
        }

        private void ListViewItem_Selected(object sender, RoutedEventArgs e)
        {

        }

        private void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            switch (PageNum)
            {
                case 0:
                    this.Insert.IsSelected = false;
                    this.Complete.IsSelected = false;
                    this.Process.IsSelected = true;
                    NextPageButton.Visibility = Visibility.Visible;
                    PreviewPageButton.Visibility = Visibility.Visible;
                    HomeButton.Visibility = Visibility.Hidden;
                    MainControl.Content = new ProgressView();
                    PageNum += 1;
                    break;
                case 1:
                    this.Insert.IsSelected = false;
                    this.Process.IsSelected = false;
                    this.Complete.IsSelected = true;
                    NextPageButton.Visibility = Visibility.Hidden;
                    PreviewPageButton.Visibility = Visibility.Visible;
                    HomeButton.Visibility = Visibility.Visible;
                    MainControl.Content = new ResultView();
                    PageNum += 1;
                    break;
                case 2:
                    NextPageButton.Visibility = Visibility.Hidden;
                    PreviewPageButton.Visibility = Visibility.Visible;
                    HomeButton.Visibility = Visibility.Visible;
                    break;
            }


        }

        private void PreviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            switch (PageNum)
            {
                case 0:
                    NextPageButton.Visibility = Visibility.Visible;
                    PreviewPageButton.Visibility = Visibility.Hidden;
                    HomeButton.Visibility = Visibility.Hidden;
                    break;
                case 1:
                    this.Process.IsSelected = false;
                    this.Complete.IsSelected = false;
                    this.Insert.IsSelected = true;
                    NextPageButton.Visibility = Visibility.Visible;
                    PreviewPageButton.Visibility = Visibility.Hidden;
                    HomeButton.Visibility = Visibility.Hidden;
                    MainControl.Content = new InsertView();
                    PageNum -= 1;
                    break;
                case 2:
                    this.Insert.IsSelected = false;
                    this.Complete.IsSelected = false;
                    this.Process.IsSelected = true;
                    NextPageButton.Visibility = Visibility.Visible;
                    PreviewPageButton.Visibility = Visibility.Visible;
                    HomeButton.Visibility = Visibility.Hidden;
                    MainControl.Content = new ProgressView();
                    PageNum -= 1;
                    break;
            }
        }
        private void HomeButton_Click(object sender, RoutedEventArgs e)
        {
            this.Process.IsSelected = false;
            this.Complete.IsSelected = false;
            this.Insert.IsSelected = true;
            NextPageButton.Visibility = Visibility.Visible;
            PreviewPageButton.Visibility = Visibility.Hidden;
            HomeButton.Visibility = Visibility.Hidden;
            MainControl.Content = new InsertView();
            PageNum = 0;
        }

        private void root_Loaded(object sender, RoutedEventArgs e)
        {
            Config = Config.GetConfig();
        }
    }
}
