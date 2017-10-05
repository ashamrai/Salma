using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace WordToTFS.ViewModel.RequiredFields
{
  public class HtmlFieldViewModel:IViewModel
  {
      private string name;
      private string value;

      public HtmlFieldViewModel(string name)
      {
          this.name = name;
          value = "";
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
              this.value = value;


              if (null != this.PropertyChanged)
              {
                  PropertyChanged(this, new PropertyChangedEventArgs("Value"));
              }
          }
      }

      public bool IsNumeric {
          get { return false; }
      }
      public event PropertyChangedEventHandler PropertyChanged;
      public string GetName()
      {
          return name;
      }

        public object GetValue()
        {
            return value;
        }
    }
}
