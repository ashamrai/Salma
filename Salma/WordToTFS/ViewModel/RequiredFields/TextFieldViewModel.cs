using System.ComponentModel;
using System.ComponentModel.DataAnnotations;


namespace WordToTFS.ViewModel.RequiredFields
{
    public class TextFieldViewModel : IViewModel
    {
        private string value;
        private string name;
        private bool isNumeric;

        public TextFieldViewModel(string name,bool numeric,string value)
        {
            this.name = name;
            this.value = value;
            isNumeric = numeric;
        }

        public string Name
        {
            get { return string.Format("{0}", name); }

            set { name = value; }
        }

        public bool IsNumeric { get { return isNumeric; } }

        public string Value
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