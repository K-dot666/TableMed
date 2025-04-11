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
using System.Collections.ObjectModel;
using Microsoft.Win32;
namespace TableMed
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<string[]> data { get; set; } = new ObservableCollection<string[]>();
        public ObservableCollection<string[]> dataTemp { get; set; } = new ObservableCollection<string[]>();
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }
        private void Search_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(District.Text)|| string.IsNullOrEmpty(BirthDate.Text)||
                string.IsNullOrEmpty(MidName.Text)|| string.IsNullOrEmpty(LastName.Text)|| string.IsNullOrEmpty(FirstName.Text))
            {
                return;
            }
            else
            {
                TableM.ItemsSource=null;
                var SearchDateB =BirthDate.Text.ToLower();
                var SearchDist = District.Text.ToLower();
                var SearchMidN = MidName.Text.ToLower();
                var SearchFirstN = FirstName.Text.ToLower();
                var SearchLastN = LastName.Text.ToLower();
                foreach (var item in data)
                {
                    if(item.ToString().ToLower().Contains(SearchDateB)||item.ToString().ToLower().Contains(SearchLastN)||
                        item.ToString().ToLower().Contains(SearchMidN)|| item.ToString().ToLower().Contains(SearchFirstN)||
                        item.ToString().ToLower().Contains(SearchDist))
                    {
                        dataTemp.Add(item);
                        TableM.ItemsSource = dataTemp;
                    }
                }
            }
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {

        }
        private void Load_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xlsx)|*.xlsx";

            try
            {
                if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.FileName))
                {
                    TableM.Columns.Clear();
                    data.Clear();

                    using (var workbook = new XLWorkbook(dlg.FileName))
                    {
                        var sheets = workbook.Worksheets.ToList();
                        if (!sheets.Any())
                            throw new InvalidOperationException("Файл не содержит листов.");

                        var sheet = sheets[0];
                        var headers = sheet.FirstRowUsed();

                        // Создаем список ожидаемых колонок
                        var requiredColumns = new List<string>
                        {
                            "Фамилия",
                            "Имя",
                            "Отчество",
                            "Дата рождения",
                            "Район"
                        };

                        // Находим все заголовки колонок
                        var actualHeaders = headers.Cells()
                            .Where(cell => !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                            .Select(cell => cell.Value.ToString())
                            .ToList();

                        // Проверяем наличие всех обязательных колонок
                        var missingColumns = requiredColumns
                            .Except(actualHeaders)
                            .ToList();

                        if (missingColumns.Any())
                        {
                            var errorMessage = $"Файл не содержит следующие обязательные столбцы:\n{string.Join("\n", missingColumns)}";
                            MessageBox.Show(errorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        // Создаем колонки таблицы
                        for (int i = 0; i < actualHeaders.Count; i++)
                        {
                            var header = actualHeaders[i];
                            var column = new DataGridTextColumn
                            {
                                Header = header,
                                Binding = new Binding($"[{i}]")
                            };
                            TableM.Columns.Add(column);
                        }

                        // Читаем данные
                        var datarow = headers.RowBelow();
                        while (!datarow.IsEmpty())
                        {
                            var rowValues = datarow.Cells()
                                .Select(c => c.Value.ToString() ?? "")
                                .ToArray();

                            if (rowValues.Any(v => !string.IsNullOrWhiteSpace(v)))
                            {
                                data.Add(rowValues);
                            }

                            datarow = datarow.RowBelow();
                        }
                    }

                    TableM.ItemsSource = null;
                    TableM.ItemsSource = data;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при чтении файла Excel",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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
