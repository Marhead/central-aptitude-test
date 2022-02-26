using CentralAptitudeTest.Models;
using CentralAptitudeTest.Commands;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows;
using System;
using Microsoft.Office.Interop.Excel;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ProgressView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProgressView : UserControl
    {
        private Config Config;
        private List<Dictionary<string, List<string>>> Temp_College_Dictionarys;
        //private ExcelManipulation ExcelManipulation = new ExcelManipulation();

        public ProgressView()
        {
            InitializeComponent();
            Config = Config.GetConfig();
        }

        private void UploadButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Config conf = new Config();

            conf.FilePath = new FilePath()
            {
                filePath = Config.FilePath.filePath,
                College_Dictionarys = Temp_College_Dictionarys,
            };
            Config.SetConfig(conf);
            return;
        }

        private void AddCollegeButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 단과대 정보 추가
            if (Config.FilePath != null)
            {
                Config conf = new Config();

                conf.Subjects = new List<string> { subject1.Text, subject2.Text, subject3.Text, subject4.Text, subject5.Text, subject6.Text };
                Dictionary<string, List<string>> dictionary = new Dictionary<string, List<string>>() {
                        { college.Text, conf.Subjects },
                    };

                if (Temp_College_Dictionarys != null)
                {
                    Temp_College_Dictionarys.Add(dictionary);
                }
                else
                {
                    Temp_College_Dictionarys = new List<Dictionary<string, List<string>>>() { { dictionary }, };
                }

                foreach (string key in dictionary.Keys)
                {
                    this.college_combo.Items.Add(key);
                }
                return;
            }
        }

        private void Input_Complete_Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // ExcelManipulation 함수 호출

            ExcelManipulation excelManipulation = new ExcelManipulation("C:\\Test\\ExcelTest\test.xlsx");

            //try
            //{
            //    ExcelManipulation excelManipulation = new ExcelManipulation(Config);
            //} catch(System.Exception exception)
            //{
            //    Console.WriteLine(exception);
            //}
        }
    }
}
