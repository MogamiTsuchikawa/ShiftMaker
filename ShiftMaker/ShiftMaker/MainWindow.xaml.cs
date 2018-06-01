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

namespace ShiftMaker
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string[] test_name = { "A","B","C","D","E","F"};
            int x = 1, y = 1;
            while (x < 13)
            {
                y = 1;
                while (y < 27)
                {
                    string combo_name = "C" + x.ToString() + "_" + y.ToString();
                    ComboBox target_combo = this.FindName(combo_name) as ComboBox;
                    if(target_combo != null)
                    {
                        foreach(string name in test_name)target_combo.Items.Add(name);
                    }
                    ++y;
                }
                ++x;
            }
        }
    }
}
