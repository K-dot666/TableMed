using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity.Core.Metadata.Edm;
using System.IO.Packaging;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using System.Windows.Controls;
using System.Windows;
using System.Xml.Linq;

namespace TableMed
{
    public class TableItem : INotifyPropertyChanged
    {
        private string[] _data;

        public TableItem(string[] data)
        {
            _data = data;
        }

        public string this[int index]
        {
            get => _data[index];
            set
            {
                if (_data[index] != value)
                {
                    _data[index] = value;
                    OnPropertyChanged(new PropertyChangedEventArgs(nameof(ItemString)));
                    OnPropertyChanged(new PropertyChangedEventArgs($"[{index}]"));
                }
            }
        }

        public string ItemString => string.Join(";", _data);

        public string[] GetData() => _data;

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
    }
}
