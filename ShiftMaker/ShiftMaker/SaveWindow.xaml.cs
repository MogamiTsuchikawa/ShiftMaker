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
using System.Windows.Shapes;

namespace ShiftMaker
{
    /// <summary>
    /// SaveWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class SaveWindow : Window
    {
        DateTime monday = DateTime.Now;
        public SaveWindow()
        {
            InitializeComponent();
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.monday = monday;
            this.Close();
        }

        private void BackWeek_Click(object sender, RoutedEventArgs e)
        {
            TimeSpan aWeek = new TimeSpan(7, 0, 0, 0);
            monday -= aWeek;
            DateLabel.Content = monday.ToShortDateString();
        }

        private void NextWeek_Click(object sender, RoutedEventArgs e)
        {
            TimeSpan aWeek = new TimeSpan(7, 0, 0, 0);
            monday += aWeek;
            DateLabel.Content = monday.ToShortDateString();
        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            if(MainWindow.monday == new DateTime(1900, 1, 1))
            {
                var today = DateTime.Parse(DateTime.Now.Year.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString());
                var TodayOfWeek = today.DayOfWeek;
                switch (TodayOfWeek)
                {
                    case DayOfWeek.Tuesday:
                        today += new TimeSpan(6, 0, 0, 0);
                        break;
                    case DayOfWeek.Wednesday:
                        today += new TimeSpan(5, 0, 0, 0);
                        break;
                    case DayOfWeek.Thursday:
                        today += new TimeSpan(4, 0, 0, 0);
                        break;
                    case DayOfWeek.Friday:
                        today += new TimeSpan(3, 0, 0, 0);
                        break;
                    case DayOfWeek.Saturday:
                        today += new TimeSpan(2, 0, 0, 0);
                        break;
                    case DayOfWeek.Sunday:
                        today += new TimeSpan(1, 0, 0, 0);
                        break;
                }
                monday = today;
                DateLabel.Content = monday.ToShortDateString();
            }
            else
            {
                monday = MainWindow.monday;
                DateLabel.Content = monday.ToShortDateString();
            }
            
        }
    }
}
