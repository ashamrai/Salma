using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System;


namespace WordToTFS.ViewModel.RequiredFields
{
    public class DateTimeFieldViewModel : IViewModel
    {
        private DateTime value;
        private string name;


        public DateTimeFieldViewModel(string name)
        {
            this.name = name;
            this.value = DateTime.Now;
        }

        public string Name
        {
            get { return string.Format("{0}", name); }

            set { name = value; }
        }

        public bool IsNumeric
        {
            get { return false; }
        }

        public DateTime Value
        {
            get { return value; }

            set
            {
                this.value = value;
                
                
                if (null != this.PropertyChanged)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Value"));
                }
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        public string GetName()
        {
            return name;
        }

        public object GetValue()
        {
            return Value;
        }
    }
}