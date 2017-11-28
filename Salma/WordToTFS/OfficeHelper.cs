using System;
using System.IO;
using WordToTFS.ConfigHelpers;
using Ionic.Zip;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace WordToTFS
{
    

    public static class OfficeHelper
    {

        public static Range CreateParagraphRange(ref Document doc)
        {
            var p = doc.Paragraphs.Add();
            p.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            return p.Range;
        }

        internal static void SetText(ref Document doc, string text, Object style)
        {
            Object styleTitle = style;
            var pt = CreateParagraphRange(ref doc);
            pt.Text = text;
            pt.set_Style(ref styleTitle);
            doc.Content.InsertParagraphAfter();
        }

        internal static void CompressDirectory(string sInDir, string sOutFile)
        {
            string[] sFiles = Directory.GetFiles(sInDir, "*.*", SearchOption.AllDirectories);
            using (var zip = new ZipFile())
            {
                foreach (string sFilePath in sFiles)
                {
                    zip.AddFile(sFilePath, string.Empty);
                }

                zip.Save(sOutFile);
            }
        }

        /// <summary>
        /// Indicates what MS Word version is used
        /// </summary>
        /// <param name="applicationVersion"></param>
        /// <returns></returns>
        public static MsWordVersion GetMsWordVersion(string applicationVersion)
        {
            switch (applicationVersion)
            {
                case "12.0":
                    return MsWordVersion.MsWord2007;
                case "14.0":
                    return MsWordVersion.MsWord2010;
                case "15.0":
                    return MsWordVersion.MsWord2013;
                case "16.0":
                    return MsWordVersion.MsWord2016;
                default:
                    return MsWordVersion.UnknownVersion;
            }
        }


        public static string GetFileType(FileInfo fileInfo, bool returnDescription)
        {
            if (fileInfo == null)
            {
                throw new ArgumentNullException("fileInfo");
            }

            string description = "File";
            if (string.IsNullOrEmpty(fileInfo.Extension))
            {
                return description;
            }

            description = string.Format("{0} File", fileInfo.Extension.Substring(1).ToUpper());
            RegistryKey typeKey = Registry.ClassesRoot.OpenSubKey(fileInfo.Extension);
            if (typeKey == null)
            {
                return description;
            }
            string type = Convert.ToString(typeKey.GetValue(string.Empty));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(type);
            if (key == null)
            {
                return description;
            }

            if (returnDescription)
            {
                description = Convert.ToString(key.GetValue(string.Empty));
                return description;
            }

            return type;
        }

        internal static void ExtractIcon(string extension, out string iconPath, out int iconIndex)
        {
            if (extension[0] != '.')
                extension = '.' + extension;

            //opens the registry for the wanted key.
            RegistryKey Root = Registry.ClassesRoot;
            RegistryKey ExtensionKey = Root.OpenSubKey(extension);
            ExtensionKey.GetValueNames();
            RegistryKey ApplicationKey = Root.OpenSubKey(ExtensionKey.GetValue("").ToString());

            //gets the name of the file that have the icon.
            string IconLocation = ApplicationKey.OpenSubKey("DefaultIcon").GetValue("").ToString();
            string[] IconPath = IconLocation.Split(',');

            if (IconPath.Length >= 2 && IconPath[1] == null)
                IconPath[1] = "0";

            iconPath = IconPath[0];

            if (IconPath.Length >= 2)
                iconIndex = int.Parse(IconPath[1]);
            else
                iconIndex = 0;
        }

        /// <summary>
        /// Get icon.
        /// </summary>
        /// <param name="icon">
        /// Word Icon.
        /// </param>
        /// <returns>
        /// msoid.
        /// </returns>
        public static string GetImageMso(Icons icon, MsWordVersion version)
        {            
            return SectionManager.Section.Images[icon.ToString()].Value;
        }
    }

    /// <summary>
    /// Word Icon
    /// </summary>
    public enum Icons
    {
        Report,
        OpenWorkItem,
        AddNewWorkItem,
        Update,
        ExportItem,
        ImportItems,
        AddDetails,
        TraceabilityMatrix,
        LinkItems,
        Help,
        Activate,
        Connect,
        Disconnect,
        MyQueries,
        TeamQueries,
        Folder,
        FlatView,
        DirectView,
        HierarchicalView,
        btnHelp,
        SyncConnectedTool,
        ShowCommentsMenu,
        groupManageWI,
        groupReporting,
        ExpandWorkItem,
        CollapseWorkItem,
        ShowPanel,
        ObsoleteWorkItem,
        Settings
    }

    /// <summary>
    /// The ms word version.
    /// </summary>
    public enum MsWordVersion
    {
        MsWord2007,
        MsWord2010,
        MsWord2013,
        MsWord2016,
        UnknownVersion
    }

}
