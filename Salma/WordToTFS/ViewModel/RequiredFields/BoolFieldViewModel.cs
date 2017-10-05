using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace WordToTFS.ViewModel.RequiredFields
{
   public class BoolFieldViewModel:IViewModel
   {
       private string name;
        public List<bool> ComboBoxCollection { get; set; }
        private bool value;

        public BoolFieldViewModel(string name, List<bool> boolResult)
        {
            this.name = name;
            this.ComboBoxCollection = boolResult;
            value = false;
        }

       public string Name
       {
           get { return string.Format("{0}", name); }
           set { name = value; }
       }

       public bool Value
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
            return Value.ToString();
        }

       public bool IsNumeric
       {
           get { return false; }
       }


   }
}
