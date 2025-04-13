using ClosedXML.Excel;
using Microsoft.Win32;
using Microsoft.Xaml.Behaviors.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
namespace TableMed
{
    public partial class MainWindow : Window
    {
<<<<<<< HEAD
        public ObservableCollection<Person> data = new ObservableCollection<Person>();
=======
        public ObservableCollection<Person> data=new ObservableCollection<Person>();
>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7
        public ObservableCollection<Person> dataTemp = new ObservableCollection<Person>();
        public DateTime date = new DateTime();
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
                try
                {
                    if (rowIndex >= 0 && columnIndex >= 0 && rowIndex < data.Count)
                    {
                        Person person = e.Row.Item as Person;
                        if (person != null)
                        {
                            string columnName = e.Column.Header as string;
                            string newValue = e.EditingElement.ToString() ?? "";

                            if (columnName == "Дата рождения")
                            {

                                if (!DateTime.TryParse(newValue, out DateTime date)|| Regex.IsMatch(newValue, @"^\d{2}\.\d{2}\.\d{4}$")==false)
                                {
                                    MessageBox.Show("Неверный формат даты. Используйте формат ДД.ММ.ГГГГ",
                                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                    e.Cancel = true;
                                    return;
                                }
                            }

                            switch (columnName)
                            {
                                case "Фамилия":
                                    person.Фамилия = newValue;
                                    break;
                                case "Имя":
                                    person.Имя = newValue;
                                    break;
                                case "Отчество":
                                    person.Отчество = newValue;
                                    break;
                                case "Дата рождения":
                                    person.Дата_рождения = date;
                                    break;
                                case "Район":
                                    person.Район = newValue;
                                    break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentFilePath))
            {
                MessageBox.Show("Файл не выбран", "Ошибка",MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                using (var workbook = new XLWorkbook(CurrentFilePath))
                {
                    var worksheet = workbook.Worksheets.Worksheet(1);
                    var headerStyle = worksheet.Row(1).Style;
                    worksheet.Range(worksheet.Row(2).FirstCell().Address,
                        worksheet.LastCell().Address).Clear();

                    // Устанавливаем формат даты для всей колонки с датами
                    worksheet.Column(4).Style.NumberFormat.Format = "dd.mm.yyyy";

                    for (int i = 0; i < data.Count; i++)
                    {
                        var person = data[i];
                        worksheet.Cell(i + 2, 1).Value = person.Фамилия;
                        worksheet.Cell(i + 2, 2).Value = person.Имя;
                        worksheet.Cell(i + 2, 3).Value = person.Отчество;
                        worksheet.Cell(i + 2, 4).Value = person.Дата_рождения.ToString("d");
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
        private void Load_Click(object sender, RoutedEventArgs e)
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
<<<<<<< HEAD
=======

                        // Создаем список ожидаемых колонок
                        var requiredColumns = new List<string>
                {
                    "Фамилия",
                    "Имя",
                    "Отчество",
                    "Дата_рождения",
                    "Район"
                };
>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7

                        var requiredColumns = new List<string> { "Фамилия", "Имя", "Отчество", "Дата_рождения", "Район" };
                        var actualHeaders = headers.Cells()
                            .Where(cell => !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                            .Select(cell => cell.Value.ToString())
                            .ToList();

                        var missingColumns = requiredColumns.Except(actualHeaders).ToList();
                        if (missingColumns.Any())
                        {
                            var errorMessage = $"Файл не содержит следующие обязательные столбцы:\n{string.Join("\n", missingColumns)}";
                            MessageBox.Show(errorMessage, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

<<<<<<< HEAD
=======
                        // Создаем колонки DataGrid с правильным binding'ом
>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7
                        foreach (var header in requiredColumns)
                        {
                            TableM.Columns.Add(new DataGridTextColumn
                            {
                                Header = header,
                                Binding = new Binding(header)
                            });
                            TableM.UpdateLayout();
                        }

                        var datarow = headers.RowBelow();
<<<<<<< HEAD
                        while (!datarow.IsEmpty())
                        {
                            var rowValues = datarow.Cells().Select(c =>
                            {
                                string value = c.Value.ToString() ?? "";
                                return value.Trim();
                            }).ToArray();

                            bool isValidRow = true;
                            if (!string.IsNullOrWhiteSpace(rowValues[actualHeaders.IndexOf("Дата_рождения")]))
                            {
                                if (!DateTime.TryParse(rowValues[actualHeaders.IndexOf("Дата_рождения")], out _))
                                {
                                    isValidRow = false;
                                }
                            }

                            if (isValidRow)
                            {
=======
                        int rowCount = 0;
                        int validRowCount = 0;

                        // Добавляем отладочный вывод для каждой строки
                        while (!datarow.IsEmpty())
                        {
                            rowCount++;
                            var rowValues = datarow.Cells().Select(c =>
                            {
                                string value = c.Value.ToString() ?? "";
                                return value.Trim(); // Удаляем пробелы
                            }).ToArray();
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

                            Debug.WriteLine($"Строка {rowCount}:");
                            Debug.WriteLine($"Значения: [{string.Join(" | ", rowValues)}]");

                            if (isValidRow)
                            {
                                validRowCount++;
                                // Добавьте перед созданием объекта Person
                                Debug.WriteLine("Индексы заголовков:");
                                foreach (var header in actualHeaders)
                                {
                                    Debug.WriteLine($"  {header}: {actualHeaders.IndexOf(header)}");
                                }
>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7
                                var person = new Person
                                {
                                    Фамилия = rowValues[actualHeaders.IndexOf("Фамилия")],
                                    Имя = rowValues[actualHeaders.IndexOf("Имя")],
                                    Отчество = rowValues[actualHeaders.IndexOf("Отчество")],
<<<<<<< HEAD
                                    Дата_рождения = DateTime.Parse(rowValues[actualHeaders.IndexOf("Дата_рождения")]),
                                    Район = rowValues[actualHeaders.IndexOf("Район")]
                                };
=======
                                    Дата_рождения = rowValues[actualHeaders.IndexOf("Дата_рождения")],
                                    Район = rowValues[actualHeaders.IndexOf("Район")]
                                };
                                Debug.WriteLine($"Добавлена строка {validRowCount}:");
                                foreach (var prop in typeof(Person).GetProperties())
                                {
                                    Debug.WriteLine($"{prop.Name}: {prop.GetValue(person)}");
                                }

>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7
                                data.Add(person);
                            }
                            datarow = datarow.RowBelow();
                        }
<<<<<<< HEAD
                        TableM.ItemsSource = Data;
                        TableM.UpdateLayout();
=======

                        Debug.WriteLine($"Обработано строк: {rowCount}");
                        Debug.WriteLine($"Валидных строк добавлено: {validRowCount}");
                        Debug.WriteLine($"Количество элементов в dataTemp: {dataTemp.Count}");
                        Debug.WriteLine($"Количество элементов в Data: {Data.Count}");
>>>>>>> 348de83d65cb616e89687db044c0573db611d2d7
                    }
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
            if (string.IsNullOrEmpty(District.Text) || string.IsNullOrEmpty(BirthDate.Text) ||
                string.IsNullOrEmpty(MidName.Text) || string.IsNullOrEmpty(LastName.Text) || string.IsNullOrEmpty(FirstName.Text))
            {
                return;
            }
            else
            {
                TableM.ItemsSource = null;
                var SearchDateB = BirthDate.Text.ToLower();
                var SearchDist = District.Text.ToLower();
                var SearchMidN = MidName.Text.ToLower();
                var SearchFirstN = FirstName.Text.ToLower();
                var SearchLastN = LastName.Text.ToLower();
                foreach (var item in data)
                {
                    if (item.ToString().ToLower().Contains(SearchDateB) || item.ToString().ToLower().Contains(SearchLastN) ||
                        item.ToString().ToLower().Contains(SearchMidN) || item.ToString().ToLower().Contains(SearchFirstN) ||
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
            string date = BirthDate.Text;
            if (!regex.IsMatch(date))
            {
                BirthDate.BorderBrush = Brushes.Red;
            }
        }
        private bool IsValidDate(string dateString)
        {
            return Regex.IsMatch(dateString, @"^\d{2}\.\d{2}\.\d{4}$") &&
                   DateTime.TryParse(dateString, out _);
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