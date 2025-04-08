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
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
namespace TableMed
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Search_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Load_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            if(dlg.ShowDialog()==true && !string.IsNullOrWhiteSpace(dlg.FileName))
            {
                TableM.ItemsSource = dlg.FileName;
            }
        }
        private void BirthDate_TextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex(@"\d{2}/./\d{2}/./\d{4}");
            string date=BirthDate.Text;
            if (!regex.IsMatch(date))
            {
                BirthDate.BorderBrush=Brushes.Red;
            }
        }
    }
}
