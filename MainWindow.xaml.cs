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
        public ObservableCollection<Person> data = new ObservableCollection<Person>();
        public ObservableCollection<Person> dataTemp = new ObservableCollection<Person>();
        public DateTime date = new DateTime();
        private string currentFilePath;
        public MainWindow()
        {
            InitializeComponent();
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

                            if (columnName == "Дата рождения")
                            {

                                if (!Regex.IsMatch(e.EditingElement.ToString(), @"^(0[1-9]|[12][0-9]|3[01])[-.](0[1-9]|1[0-2])[-.]\d{4}$"))
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
                                    person.Фамилия = e.EditingElement.ToString() ?? "";
                                    break;
                                case "Имя":
                                    person.Имя = e.EditingElement.ToString() ?? "";
                                    break;
                                case "Отчество":
                                    person.Отчество = e.EditingElement.ToString() ?? "";
                                    break;
                                case "Дата рождения":
                                    person.Дата_рождения = DateTime.Parse(e.EditingElement.ToString() ?? "");
                                    break;
                                case "Район":
                                    person.Район = e.EditingElement.ToString() ?? "";
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
                MessageBox.Show("Файл не выбран", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try
            {
                using (var workbook = new XLWorkbook(CurrentFilePath))
                {
                    var worksheet = workbook.Worksheets.Worksheet(1);
                    worksheet.Range(worksheet.Row(2).FirstCell().Address,worksheet.LastCell().Address).Clear();                    
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
                    MessageBox.Show("Изменения успешно сохранены", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                                var person = new Person
                                {
                                    Фамилия = rowValues[actualHeaders.IndexOf("Фамилия")],
                                    Имя = rowValues[actualHeaders.IndexOf("Имя")],
                                    Отчество = rowValues[actualHeaders.IndexOf("Отчество")],
                                    Дата_рождения = DateTime.Parse(rowValues[actualHeaders.IndexOf("Дата_рождения")]),
                                    Район = rowValues[actualHeaders.IndexOf("Район")]
                                };
                                data.Add(person);
                            }
                            datarow = datarow.RowBelow();
                        }
                        TableM.ItemsSource = Data;
                        TableM.UpdateLayout();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при чтении файла Excel",MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Search_Click(object sender, RoutedEventArgs e)
        {
            // Очистка предыдущих результатов
            dataTemp.Clear();
            TableM.ItemsSource = null;
            // Получение значений из полей поиска
            var searchLastName = LastName.Text.ToLower();
            var searchFirstName = FirstName.Text.ToLower();
            var searchMidName = MidName.Text.ToLower();
            var searchDistrict = District.Text.ToLower();
            var searchBirthDate = BirthDate.Text.ToLower();
            Regex dateRegex = new Regex(@"^(0[1-9]|[12][0-9]|3[01])[-.](0[1-9]|1[0-2])[-.]\d{4}$");
            if (!string.IsNullOrEmpty(searchBirthDate) && !dateRegex.IsMatch(searchBirthDate))
            {
                BirthDate.BorderBrush = Brushes.Red;
                return;
            }
            //поиск совпадений по каждому полю
            foreach (var person in Data)
            {
                bool isMatch = true;
                if (!string.IsNullOrEmpty(searchLastName) &&!person.Фамилия.ToLower().Contains(searchLastName))
                {
                    isMatch = false;
                }
                if (!string.IsNullOrEmpty(searchFirstName) &&!person.Имя.ToLower().Contains(searchFirstName))
                {
                    isMatch = false;
                }
                if (!string.IsNullOrEmpty(searchMidName) && !person.Отчество.ToLower().Contains(searchMidName))
                {
                    isMatch = false;
                }
                if (!string.IsNullOrEmpty(searchDistrict) &&!person.Район.ToLower().Contains(searchDistrict))
                {
                    isMatch = false;
                }
                if (!string.IsNullOrEmpty(searchBirthDate))
                {
                    string dateStr = person.Дата_рождения.ToString("dd.MM.yyyy");
                    if (!dateStr.Contains(searchBirthDate))
                    {
                        isMatch = false;
                    }
                }
                // Добавление совпадения
                if (isMatch)
                {
                    dataTemp.Add(person);
                }
            }
            // рефреш 
            if (dataTemp.Count > 0)
            {
                TableM.ItemsSource = dataTemp;
                TableM.UpdateLayout();
            }
            else
            {
                MessageBox.Show("Ничего не найдено", "Результат поиска",MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            LastName.Text = string.Empty;
            FirstName.Text = string.Empty;
            MidName.Text = string.Empty;
            District.Text = string.Empty;
            BirthDate.Text = string.Empty;
            TableM.SelectedIndex = -1;
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