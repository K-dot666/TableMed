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
using System.ComponentModel;
using System.Diagnostics;
namespace TableMed
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<Person> data=new ObservableCollection<Person>();
        public ObservableCollection<Person> dataTemp = new ObservableCollection<Person>();
        private string currentFilePath;
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            TableM.ItemsSource = Data;

        }
        public ObservableCollection<Person> Data
        {
            get => data;
            set
            {
                data = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Data)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
        public string CurrentFilePath
        {
            get => currentFilePath;
            set
            {
                currentFilePath = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(CurrentFilePath)));
            }
        }
        private void TableM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var rowIndex = e.Row.GetIndex();
                var columnIndex = e.Column.DisplayIndex;

                if (rowIndex >= 0 && columnIndex >= 0 && rowIndex < data.Count)
                {
                    Person person = e.Row.Item as Person;
                    if (person != null)
                    {
                        // Определяем изменённое свойство
                        string columnName = e.Column.Header as string;

                        // Обновляем соответствующее свойство
                        switch (columnName)
                        {
                            case "Фамилия":
                                person.Фамилия = e.EditingElement.ToString() ?? "";
                                break;
                            case "Имя":
                                person.Имя = e.EditingElement.ToString() ?? "";
                                break;
                            case "Отчество":
                                person.Отчество = e.EditingElement.ToString() ?? "";
                                break;
                            case "Дата рождения":
                                person.Дата_рождения = e.EditingElement.ToString() ?? "";
                                break;
                            case "Район":
                                person.Район = e.EditingElement.ToString() ?? "";
                                break;
                        }
                    }
                }
            }
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentFilePath))
            {
                MessageBox.Show("Файл не выбран", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                using (var workbook = new XLWorkbook(CurrentFilePath))
                {
                    var worksheet = workbook.Worksheets.Worksheet(1);
                    // Сохраняем форматирование заголовков
                    var headerStyle = worksheet.Row(1).Style;
                    // Очищаем только данные, оставляя заголовки
                    worksheet.Range(worksheet.Row(2).FirstCell().Address,worksheet.LastCell().Address).Clear();

                    // Записываем данные, начиная со второй строки
                    for (int i = 0; i < data.Count; i++)
                    {
                        var person = data[i];
                        worksheet.Cell(i + 2, 1).Value = person.Фамилия;
                        worksheet.Cell(i + 2, 2).Value = person.Имя;
                        worksheet.Cell(i + 2, 3).Value = person.Отчество;
                        worksheet.Cell(i + 2, 4).Value = person.Дата_рождения;
                        worksheet.Cell(i + 2, 5).Value = person.Район;
                    }
                    workbook.Save();
                    MessageBox.Show("Изменения успешно сохранены","Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}","Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private  void Load_Click(object sender, RoutedEventArgs e)
        {
            Data.Clear();
            TableM.UpdateLayout();
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xlsx)|*.xlsx";
            try
            {
                if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.FileName))
                {
                    CurrentFilePath = dlg.FileName;
                    TableM.Columns.Clear();
                    dataTemp.Clear();
                        using (var workbook = new XLWorkbook(dlg.FileName))
                        {
                            var sheets = workbook.Worksheets.ToList();
                            if (!sheets.Any())
                                throw new InvalidOperationException("Файл не содержит листов.");
                            var sheet = sheets[0];
                            var headers = sheet.FirstRowUsed();

                            //  список ожидаемых заголовков
                            var requiredColumns = new List<string>{"Фамилия","Имя", "Отчество", "Дата_рождения", "Район"};
                            //  все заголовки файла
                            var actualHeaders = headers.Cells().Where(cell => !string.IsNullOrWhiteSpace(cell.Value.ToString())).Select(cell => cell.Value.ToString()).ToList();
                            // проверка заголовков
                            var missingColumns = requiredColumns.Except(actualHeaders).ToList();
                            if (missingColumns.Any())
                            {
                                var errorMessage = $"Файл не содержит следующие обязательные столбцы:\n{string.Join("\n", missingColumns)}";MessageBox.Show(errorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            // Создание заголовков
                            foreach (var header in requiredColumns)
                            {
                                TableM.Columns.Add(new DataGridTextColumn
                                {
                                    Header = header,
                                    Binding = new Binding(header),
                                    IsReadOnly = false
                                });
                                TableM.UpdateLayout();
                            }
                            // Читаем данные и отсеиваем пустые строки
                            var datarow = headers.RowBelow();
                            while (!datarow.IsEmpty())
                            {
                                var rowValues = datarow.Cells().Select(c =>
                                { string value = c.Value.ToString() ?? "";return value.Trim(); }).ToArray();
                                bool isValidRow = false;
                                foreach (var column in requiredColumns)
                                {
                                    int columnIndex = actualHeaders.IndexOf(column);
                                    if (!string.IsNullOrWhiteSpace(rowValues[columnIndex]))
                                    {
                                        isValidRow = true;
                                        break;
                                    }
                                }
                                if (isValidRow)
                                {
                                    // Добавление записи объекта Person в коллекцию
                                    var person = new Person
                                    {
                                        Фамилия = rowValues[actualHeaders.IndexOf("Фамилия")],
                                        Имя = rowValues[actualHeaders.IndexOf("Имя")],
                                        Отчество = rowValues[actualHeaders.IndexOf("Отчество")],
                                        Дата_рождения = rowValues[actualHeaders.IndexOf("Дата_рождения")],
                                        Район = rowValues[actualHeaders.IndexOf("Район")]
                                    };
                                    data.Add(person);
                                }
                                datarow = datarow.RowBelow();
                            }
                        }
                        TableM.ItemsSource = Data;
                        TableM.UpdateLayout();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при чтении файла Excel",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
/*
 * Пользователь    
  ▼  
[Загрузка файла] → [Парсинг и валидация] → [Временное хранение данных]  
  │  
  ▼  
[Ввод критериев] → [Коррекция ошибок ввода] → [Поиск по таблице]  
  │                              │  
  ▼                              ▼  
[Отображение результатов]    [Оповещение об ошибках]
 */