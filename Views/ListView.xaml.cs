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

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ListView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ListView : UserControl
    {
        // 작업 현황 글자 색
        public static readonly DependencyProperty ColorOfForegroundPro =
            DependencyProperty.Register("ColorOfForeground", typeof(Brush), typeof(ListView),
                new PropertyMetadata(Brushes.Black));

        public Brush ColorOfForeground
        {
            get { return (Brush)GetValue(ColorOfForegroundPro); }
            set { SetValue(ColorOfForegroundPro, value); }
        }

        // 작업 현황 배경 색
        public static readonly DependencyProperty ColorOfBackgroundPro =
            DependencyProperty.Register("ColorOfBackground", typeof(Brush), typeof(ListView),
                new PropertyMetadata(Brushes.Transparent));

        public Brush ColorOfBackground
        {
            get { return (Brush)GetValue(ColorOfBackgroundPro); }
            set { SetValue(ColorOfBackgroundPro, value); }
        }

        // 작업 현황 항목들 가로 길이
        public static readonly DependencyProperty Task_Width_Pro =
           DependencyProperty.Register("Task_Width", typeof(int), typeof(ListView),
               new PropertyMetadata(140));

        public int Task_Width
        {
            get { return (int)GetValue(Task_Width_Pro); }
            set { SetValue(Task_Width_Pro, value); }
        }


        // 작업 현황 항목들 세로 길이
        public static readonly DependencyProperty Task_Height_Pro =
           DependencyProperty.Register("Task_Height", typeof(int), typeof(ListView),
               new PropertyMetadata(50));

        public int Task_Height
        {
            get { return (int)GetValue(Task_Height_Pro); }
            set { SetValue(Task_Height_Pro, value); }
        }

        // 작업 현황 폰트 사이즈
        public static readonly DependencyProperty Task_FontSize_Pro =
           DependencyProperty.Register("Task_FontSize", typeof(int), typeof(ListView),
               new PropertyMetadata(19));

        public int Task_FontSize
        {
            get { return (int)GetValue(Task_FontSize_Pro); }
            set { SetValue(Task_FontSize_Pro, value); }
        }

        // 작업 현황 폰트 디자인 소스
        public static readonly DependencyProperty Task_Font_Pro =
           DependencyProperty.Register("Task_Font", typeof(string), typeof(ListView),
               new PropertyMetadata("/Resources/Fonts/NanumBarunGothic.ttf"));

        public string Task_Font
        {
            get { return (string)GetValue(Task_Font_Pro); }
            set { SetValue(Task_Font_Pro, value); }
        }

        public ListView()
        {
            InitializeComponent();
        }

    }
}
