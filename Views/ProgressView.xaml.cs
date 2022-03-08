using CentralAptitudeTest.Models;
using CentralAptitudeTest.Commands;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows;
using Microsoft.Win32;
using System;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace CentralAptitudeTest.Views
{
    /// <summary>
    /// ProgressView.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ProgressView : UserControl
    {
        private Config Config;
        private List<Dictionary<string, List<string>>> TempCollegeDictionaries;
        private Dictionary<string, List<string>> dictionary;
        private ExcelManipulation ExcelManipulation;
        private List<string> tempList;
        private int Add_Subject_Number = 1;

        public ProgressView()
        {
            InitializeComponent();
            Config = Config.GetConfig();
        }

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.ShowDialog();

            ReadFilePath(openFileDialog.FileName);
        }

        private void ReadFilePath(string filename)
        {
            myTextBox.Text = filename;
            return;
        }

        private void UploadButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Config config = new Config();

            config.FilePath = new FilePath() { whole_data_filePath = Config.FilePath.whole_data_filePath, process_data_filePath = myTextBox.Text };

            //config.CollegeDictionaries = TempCollegeDictionaries;

            Config.SetConfig(config);
            return;
        }

        private void Input_Complete_Button_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // ExcelManipulation 함수 호출
            ExcelManipulation = new ExcelManipulation(Config);

            ExcelManipulation.WriteToCell();

            ExcelManipulation.CloseFile();

            MessageBox.Show(Config.FilePath.whole_data_filePath);
        }

        //public void PutData()
        //{
        //    tempList = new List<string> { subject1.Text, subject2.Text, subject3.Text, subject4.Text, subject5.Text, subject6.Text };

        //    dictionary = new Dictionary<string, List<string>>() {
        //                { college.Text, tempList },
        //            };
        //}

        //private void AddCollegeButton_Click(object sender, System.Windows.RoutedEventArgs e)
        //{
        //    // 단과대 정보 추가
        //    if (Config.FilePath != null)
        //    {
        //        if (String.IsNullOrEmpty(college.Text))
        //        {
        //            MessageBox.Show("단과대를 입력하세요.");
        //            return;
        //        }

        //        //if (String.IsNullOrEmpty(subject1.Text))
        //        //{
        //        //    MessageBox.Show("학과명을 하나 이상 입력하세요.");
        //        //    return;
        //        //}

        //        PutData();

        //        if (TempCollegeDictionaries != null)
        //        {
        //            //foreach (Dictionary<string, List<string>> dic in TempCollegeDictionaries)
        //            //{
        //            //    foreach (string key in dic.Keys)
        //            //    {
        //            //        if (key == college.Text)
        //            //        {
        //            //            if (MessageBox.Show("이미 존재하는 단과대명입니다. 새로 입력한 정보로 바꾸시겠습니까?", "YesOrNo", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
        //            //            {
        //            //                dic.Remove(key);
        //            //                //this.college_combo.Items.Remove(key);
        //            //                //PutData();
        //            //                break;
        //            //            }
        //            //            else
        //            //            {
        //            //                return;
        //            //            }
        //            //        }
        //            //    }
        //            //}

        //            TempCollegeDictionaries.Add(dictionary);
        //        }
        //        else
        //        {
        //            TempCollegeDictionaries = new List<Dictionary<string, List<string>>>() { { dictionary }, };
        //        }

        //        //foreach (string key in dictionary.Keys)
        //        //{
        //        //    this.college_combo.Items.Add(key);
        //        //    sub_num1.Text = dictionary[key][0];
        //        //    sub_num2.Text = dictionary[key][1];
        //        //    sub_num3.Text = dictionary[key][2];
        //        //    sub_num4.Text = dictionary[key][3];
        //        //    sub_num5.Text = dictionary[key][4];
        //        //    sub_num6.Text = dictionary[key][5];
        //        //}

        //        MessageBox.Show("'" + college.Text + "'" + " 단과대 정보를 추가했습니다.");

        //        this.college.Text = "";
        //        this.subject1.Text = "";
        //        this.subject2.Text = "";
        //        this.subject3.Text = "";
        //        this.subject4.Text = "";
        //        this.subject5.Text = "";
        //        this.subject6.Text = "";

        //        this.Enter_info6.Visibility = Visibility.Hidden;
        //        this.Enter_info1.Visibility = Visibility.Visible;

        //        this.subject2.Visibility = Visibility.Hidden;
        //        this.subject3.Visibility = Visibility.Hidden;
        //        this.subject4.Visibility = Visibility.Hidden;
        //        this.subject5.Visibility = Visibility.Hidden;
        //        this.subject6.Visibility = Visibility.Hidden;

        //        Add_Subject_Number = 1;

        //        return;
        //    }
        //}

        //private void college_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    Config = Config.GetConfig();

        //    try
        //    {
        //        foreach (Dictionary<string, List<string>> dictionary in Config.CollegeDictionaries)
        //        {
        //            foreach (string key in dictionary.Keys)
        //            {
        //                if (key == college_combo.SelectedItem.ToString())
        //                {
        //                    sub_num1.Text = dictionary[key][0];
        //                    sub_num2.Text = dictionary[key][1];
        //                    sub_num3.Text = dictionary[key][2];
        //                    sub_num4.Text = dictionary[key][3];
        //                    sub_num5.Text = dictionary[key][4];
        //                    sub_num6.Text = dictionary[key][5];
        //                }
        //            }
        //        }
        //    }
        //    catch (NullReferenceException)
        //    {
        //        MessageBox.Show("단과대, 학과명을 입력 후 '데이터 올리기' 를 클릭하세요.");
        //    }
        //}


        //private void subject1_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        //{
        //    if (e.Key == System.Windows.Input.Key.Enter)
        //    {
        //        Add_Subject_Number += 1;

        //        switch (Add_Subject_Number)
        //        {
        //            case 2:
        //                this.Enter_info1.Visibility = Visibility.Hidden;
        //                this.Enter_info2.Visibility = Visibility.Visible;
        //                this.subject2.Visibility = Visibility.Visible;
        //                this.subject2.Focus();
        //                break;
        //            case 3:
        //                this.Enter_info2.Visibility = Visibility.Hidden;
        //                this.Enter_info3.Visibility = Visibility.Visible;
        //                this.subject3.Visibility = Visibility.Visible;
        //                this.subject3.Focus();
        //                break;
        //            case 4:
        //                this.Enter_info3.Visibility = Visibility.Hidden;
        //                this.Enter_info4.Visibility = Visibility.Visible;
        //                this.subject4.Visibility = Visibility.Visible;
        //                this.subject4.Focus();
        //                break;
        //            case 5:
        //                this.Enter_info4.Visibility = Visibility.Hidden;
        //                this.Enter_info5.Visibility = Visibility.Visible;
        //                this.subject5.Visibility = Visibility.Visible;
        //                this.subject5.Focus();
        //                break;
        //            case 6:
        //                this.Enter_info5.Visibility = Visibility.Hidden;
        //                this.Enter_info6.Visibility = Visibility.Visible;
        //                this.subject6.Visibility = Visibility.Visible;
        //                this.subject6.Focus();
        //                break;
        //        }
        //    }
        //}
    }
}

