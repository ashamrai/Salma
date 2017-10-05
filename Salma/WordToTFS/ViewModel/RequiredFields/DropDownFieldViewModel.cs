using System.Collections.Generic;
using System.ComponentModel;

namespace WordToTFS.ViewModel.RequiredFields
{
    /// <summary>
    /// responsible for displaying requirement field
    /// </summary>
    public class DropDownFieldViewModel : IViewModel
    {
        private string name;
        public List<string> ComboBoxCollection { get; set; }
        private string value;
        private bool isEditable;
        private bool isNumeric;

        public DropDownFieldViewModel(string name, List<string> requiredTypes,bool editable,bool numeric)
        {
            this.name = name;
            ComboBoxCollection = requiredTypes;
            isEditable = editable;
            isNumeric = numeric;
        }

        public bool IsEditable
        {
            get { return isEditable; }
        }

        public bool IsNumeric
        {
            get { return isNumeric; }
        }

        public string Name
        {
            get { return string.Format("{0}", name); }
            set { name = value; }
        }

        public string Value
        {
            get { return value; }

            set
            {
                if (null != this.PropertyChanged)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Value"));
                }
                this.value = value;
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
