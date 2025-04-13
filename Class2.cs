using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableMed
{
    public class Person : INotifyPropertyChanged
    {
        private string lastName;
        private string firstName;
        private string middleName;
        private string birthDate;
        private string district;

        public string Фамилия
        {
            get => lastName;
            set
            {
                lastName = value;
                OnPropertyChanged(nameof(Фамилия));
            }
        }

        public string Имя
        {
            get => firstName;
            set
            {
                firstName = value;
                OnPropertyChanged(nameof(Имя));
            }
        }

        public string Отчество
        {
            get => middleName;
            set
            {
                middleName = value;
                OnPropertyChanged(nameof(Отчество));
            }
        }

        public string Дата_рождения
        {
            get => birthDate;
            set
            {
                birthDate = value;
                OnPropertyChanged(nameof(Дата_рождения));
            }
        }

        public string Район
        {
            get => district;
            set
            {
                district = value;
                OnPropertyChanged(nameof(Район));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
