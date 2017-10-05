using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace WordToTFS.ConfigHelpers
{
 
    public class WordSection : ConfigurationSection
    {
        [ConfigurationProperty("images", IsDefaultCollection = true, IsRequired = true),
        ConfigurationCollection(typeof(ImageElementCollection), AddItemName = "image")]
        public ImageElementCollection Images
        {
            get
            {
                return ((ImageElementCollection)(base["images"]));
            }
        }
    }  
}
