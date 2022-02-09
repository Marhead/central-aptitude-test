using CentralAptitudeTest.Models;
using System.Windows.Controls;

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
            InitializeComponent();
            Config = Config.GetConfig();
            foreach (FilePath fp in Config.FilePaths)
            {
                if (fp.filePath != null)
                {
                    college.Content = fp.College;
                    subject.Content = fp.Subject;
                    TextBlock.Text = fp.filePath;
                    return;
                }
            }
        }
    }
}
