using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace WordToTFS.ConfigHelpers
{
    [ConfigurationCollection(typeof(ImageElement))]
    public class ImageElementCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ImageElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ImageElement)(element)).Key;
        }

        public ImageElement this[string key]
        {
            get
            {
                return (ImageElement)BaseGet(key);
            }
        }
    }
}
