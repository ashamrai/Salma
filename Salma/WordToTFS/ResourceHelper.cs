using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;

namespace WordToTFS
{
    public static class ResourceHelper
    {
        public static string GetResourceString(string resourceName)
        {
            return new ResourceManager("WordToTFS.UIResources", typeof(ResourceHelper).Assembly).GetString(resourceName);
        }
    }
}
