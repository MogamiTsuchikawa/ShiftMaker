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
using Microsoft.Win32;
using Forms = System.Windows.Forms;
using ClosedXML.Excel;

namespace ShiftMaker
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string nowPath ="";
        public static DateTime monday = new DateTime(1900,1,1);
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var name_set = new System.IO.StreamReader("name.txt");
            string[] names = name_set.ReadToEnd().Split(',');
            //string[] test_name = { "","太田","八木","山本","加藤","愛奈","一木","沖山","市川", "恵梨佳", "新村", "長谷川", "下玉利", "尚子", "宝田", "森", "野田" };
            int x = 1, y = 1;
            while (x < 13)
            {
                y = 1;
                while (y < 27)
                {
                    string combo_name = "C" + x.ToString() + "_" + y.ToString();
                    ComboBox target_combo = this.FindName(combo_name) as ComboBox;
                    //target_combo.
                    if(target_combo != null)
                    {
                        foreach(string name in names)target_combo.Items.Add(name);
                    }
                    ++y;
                }
                ++x;
            }
        }
        void SetDate(string mondayText)
        {
            monday = DateTime.Parse(mondayText);
            TimeSpan WorkDay = new TimeSpan(5, 0, 0,0);
            DateTime satday = monday + WorkDay;
            DATE.Header = monday.ToShortDateString() + "(月)～" + satday.Date.ToShortDateString()+"(土)";

        }

        void SetData(string path)
        {
            var fo = new System.IO.StreamReader(path);
            string date = fo.ReadLine();
            SetDate(date);
            string[] objects = fo.ReadLine().Split('/');
            int x = 1, y = 1,i=0;
            while (x < 13)
            {
                y = 1;
                while (y < 27)
                {
                    string combo_name = "C" + x.ToString() + "_" + y.ToString();
                    ComboBox target_combo = this.FindName(combo_name) as ComboBox;
                    target_combo.Text = objects[i];
                    ++y;
                    ++i;
                }
                ++x;
            }
            fo.Close();

        }
        void SaveData(string path)
        {
            string st_line="";
            int x = 1, y = 1, i = 0;
            while (x < 13)
            {
                y = 1;
                while (y < 27)
                {
                    string combo_name = "C" + x.ToString() + "_" + y.ToString();
                    ComboBox target_combo = this.FindName(combo_name) as ComboBox;
                    if (i == 0)
                    {
                        st_line = target_combo.Text;
                    }
                    else
                    {
                        st_line = st_line +"/"+target_combo.Text;
                    }
                    ++y;
                    ++i;
                }
                ++x;
            }
            var fs = new System.IO.StreamWriter(path);
            fs.WriteLine(monday.ToShortDateString());
            fs.Write(st_line);
            fs.Close();
        }
        private void OpenMenu_Click(object sender, RoutedEventArgs e)
        {
            // ダイアログのインスタンスを生成
            var dialog = new OpenFileDialog();

            // ファイルの種類を設定
            dialog.Filter = "テキストファイル (*.txt)|*.txt|全てのファイル (*.*)|*.*";

            // ダイアログを表示する
            if (dialog.ShowDialog() == true)
            {
                // 選択されたファイル名 (ファイルパス) をメッセージボックスに表示
                //MessageBox.Show(dialog.FileName);
                SetData(dialog.FileName);
                nowPath = dialog.FileName;
            }

        }

        private void SaveMenu_Click(object sender, RoutedEventArgs e)
        {
            if (nowPath == "")
            {
                var SaveWindow = new SaveWindow();
                SaveWindow.ShowDialog();
                var dialog = new Forms.FolderBrowserDialog();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string day = monday.ToShortDateString();
                    day = day.Replace(@"/", "-");
                    nowPath = dialog.SelectedPath + @"\シフト" + day + ".txt";
                    SaveData(nowPath);
                    
                }
            }
            else
            {
                SaveData(nowPath);
            }
        }

        private void SaveAsMenu_Click(object sender, RoutedEventArgs e)
        {
            var SaveWindow = new SaveWindow();
            SaveWindow.ShowDialog();
            var dialog = new Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string day = monday.ToShortDateString();
                day = day.Replace(@"/", "-");
                nowPath = dialog.SelectedPath + @"\シフト" + day + ".txt";
                SaveData(nowPath);
            }
        }

        private void ExportToExcelMenu_Click(object sender, RoutedEventArgs e)
        {
            if (nowPath == "")
            {
                MessageBox.Show("まずは保存をしてからにしてください！！");
            }
            else
            {
                //ExportToExcel
                createExcelFile(nowPath);
            }
        }
        void createExcelFile(string path)
        {
            
            var workbook = new XLWorkbook();
            string dayname = monday.ToShortDateString().Replace("/", "-");
            var worksheet = workbook.Worksheets.Add("シフト(" + dayname + "～)");
            worksheet.Cell(1, 1).Value = "シフト(" + monday.ToShortDateString() + "～)";
            for(int gozen=2; gozen < 13; gozen += 2)worksheet.Cell(3, gozen).Value = "午前";
            for (int gogo = 3; gogo < 14; gogo += 2) worksheet.Cell(3, gogo).Value = "午後";
            int x = 2, y = 4;
            var fo = new System.IO.StreamReader(path);
            string date = fo.ReadLine();
            SetDate(date);
            string[] objects = fo.ReadLine().Split('/');
            int i = 0;
            int Max_x = 14;

            //木曜日と土曜日の午後が必要か確認
            bool thurs_active = false, satur_active = false;
            //木曜日の午後
            int ii = 182;
            while (182 + 26 > ii)
            {
                if (objects[ii] != "") thurs_active = true;
                ++ii;
            }
            if (thurs_active)
            {
                --Max_x;
            }
            //土曜日の午後
            ii = 286;
            while (286 + 26 > ii)
            {
                if (objects[ii] != "") satur_active = false;
                ++ii;
            }
            if (satur_active)
            {
                --Max_x;
            }

            //int x = 1, y = 1, 

            while (x < Max_x)
            {
                y = 4;
                while (i<26*(x-1))
                {
                    
                    bool empty = false;
                    if (objects[i] == "")
                    {
                        empty = true;
                        int backnum = i / 26;
                        for(ii = 0; ii < backnum+1; ++ii)
                        {
                            if (objects[i - ii * 26] != "") empty = false;
                        }
                        for (ii = 0; ii < 12 - backnum;++ii) if (objects[i + ii * 26] != "") empty = false;
                    }
                    if (empty == false)
                    {
                        worksheet.Cell(y, x).Value = objects[i];
                        ++y;
                    }
                    if (x == 2)
                    {
                        switch (i)
                        {
                            case 0:
                                worksheet.Cell(y-1, x - 1).Value = "受付";
                                break;
                            case 5:
                                worksheet.Cell(y-1, x - 1).Value = "物療";
                                break;
                            case 12:
                                worksheet.Cell(y-1, x - 1).Value = "PT";
                                break;
                            case 19:
                                worksheet.Cell(y-1, x - 1).Value = "診療";
                                break;
                        }
                    }
                    ++i;
                }
                
                ++x;
            }
            fo.Close();
            path = path.Replace(".txt", "");
            path += ".xlsx";
            workbook.SaveAs(path);
            MessageBox.Show("完了");
        }
    }
}