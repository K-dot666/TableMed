using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
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
using ClosedXML;
using ClosedXML.Excel;
using ClosedXML.Parser;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
namespace TableMed
{
    public partial class MainWindow : Window
    {
        List<String> data;
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
            try
            {
                if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.FileName))
                {
                    List<String> sheets;
                    data = new List<string>();
                    sheets = new List<string>();
                    var rows = new List<List<string>>();
                    using (var workbook = new XLWorkbook(dlg.FileName))
                    {
                        foreach (IXLWorksheet worksheet in workbook.Worksheets)
                        {
                            sheets.Add(worksheet.Name);
                        }
                        var sheet = workbook.Worksheet(sheets[0]);
                        var headers = sheet.FirstRowUsed();
                        foreach (var c in headers.Cells())
                        {
                            string headertext = c.Value.ToString();
                            if (!string.IsNullOrWhiteSpace(headertext))
                            {
                                TableM.Columns.Add(new DataGridTextColumn { Header = headertext });
                            }
                        }
                        var datarow = headers.RowBelow();

                        // Читаем данные построчно
                        while (!datarow.IsEmpty())
                        {
                            var rowValues = new List<string>();
                            foreach (var v in datarow.Cells())
                            {
                                string value = v.Value.ToString();
                                rowValues.Add(value);
                            }
                            rows.Add(rowValues);
                            datarow = datarow.RowBelow();
                        }
                        TableM.ItemsSource = data;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
