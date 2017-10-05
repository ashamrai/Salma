using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace WordToTFS.ConfigHelpers
{
    public class ImageElement : ConfigurationElement
    {

        [ConfigurationProperty("key", DefaultValue = "", IsKey = true, IsRequired = true)]
        public string Key
        {
            get
            {
                return ((string)(base["key"]));
            }
            set
            {
                base["key"] = value;
            }
        }

       [ConfigurationProperty("value", DefaultValue = "", IsKey = false, IsRequired = false)]
       public string Value
       {
           get
           {
               return ((string)(base["value"]));
           }
           set
           {
               base["value"] = value;
           }
       }
    }
}
