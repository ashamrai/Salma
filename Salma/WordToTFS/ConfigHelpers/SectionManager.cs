using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Reflection;

namespace WordToTFS.ConfigHelpers
{
    public static class SectionManager
    {
        public static WordSection Section = null;

        public static void SetSection(MsWordVersion MsWordVersion)
        {                          
            Section = (MsWordVersion == MsWordVersion.MsWord2007 ?
                (WordSection)ConfigurationManager.GetSection(Properties.Resources.Word2007Section) :
                (WordSection)ConfigurationManager.GetSection(Properties.Resources.WordSection));
        }       
    }
}
