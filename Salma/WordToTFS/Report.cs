using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Controls;

using WordToTFS.View;
using MessageBox = System.Windows.MessageBox;
using WordToTFSWordAddIn.Views;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.Office.Interop.Word;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using CheckBox = System.Windows.Controls.CheckBox;
using Field = Microsoft.TeamFoundation.WorkItemTracking.Client.Field;
using System.Windows;

namespace WordToTFS
{
    using System.Text.RegularExpressions;

    using mshtml;

    using Revision = Microsoft.TeamFoundation.WorkItemTracking.Client.Revision;

    public class Report
    {
        #region Fields

        private Document document;
        private readonly TfsManager tfsManager;

        private DateTime? revisionDate;
        #endregion

        #region Constructors and Destructors

        public Report(Document activeDocument)
        {
            tfsManager = TfsManager.Instance;
            document = activeDocument;
        }

        #endregion

        #region Public Methods and Operators
        //TODO:refactor it.
        public void GenerateReport(string project, Action<string> descriptionText)
        {            
            var popup = new QueryReport { Project = project };
            Project currProject = tfsManager.GetProject(project);
            var qm = new QueryManager(currProject, popup.QueriesTree, tfsManager.ItemsStore);
            List<FieldDefinition> allFields = tfsManager.GetAllAvailableWorkItemFieldsForProject(project);
            string queryName = string.Empty;
            Query query = null;
            List<string> displayFieldList = null;

            qm.Query += qDef =>
                {
                    if (qDef != null)
                    {
                        tfsManager.ItemsStore.RefreshCache();
                        query = new Query(tfsManager.ItemsStore, qDef.QueryText);
                        displayFieldList = query.DisplayFieldList.Cast<FieldDefinition>().Select(f => f.Name).ToList();
                        popup.QueryDef = qDef;

                        // clear controls
                        popup.PropertiesList.SelectionMode = SelectionMode.Multiple;
                        popup.tbxHeader.Clear();
                        popup.PropertiesList.Items.Clear();
                        popup.PropertiesListBody.Items.Clear();
                        queryName = qDef.Name;
                        if (displayFieldList != null)
                        {
                            popup.PopulateQueryFields(allFields, displayFieldList);
                            popup.InsertButton.IsEnabled = true;
                        }
                    }
                    else
                    {
                        popup.InsertButton.IsEnabled = false;
                    }
                };



            switch (tfsManager.tfsVersion)
            {
                case TfsVersion.Tfs2011:
                    popup.ManageQueriesLink.NavigateUri = new Uri(string.Format("{0}/{1}/_workitems",tfsManager.ItemsStore.TeamProjectCollection.Uri, project));
                    break;
                case TfsVersion.Tfs2010:
                    popup.ManageQueriesLink.NavigateUri = new Uri(string.Format("{0}/web/UI/Pages/WorkItems/QueryExplorer.aspx?pguid={1}", tfsManager.GetTfsUrl(), tfsManager.GetProject(project).Guid));
                    break;
            }
            popup.Create(null, Icons.Report);
           
            //Generate Report from WI list
            if (popup.byWIIDsTabItem.IsSelected && !popup.IsCanceled)
            {
                tfsManager.ItemsStore.RefreshCache();
                query = new Query(tfsManager.ItemsStore, popup.queryDef.QueryText);
                queryName = popup.queryDef.Name;
            }

            revisionDate = popup.onDate.SelectedDate;
            if (!popup.IsCanceled)
            {
                List<string> metaDataSelectedItems = GetSelectedItemsFromCheckListBox(popup.PropertiesList);
                List<string> bodySelectedItems = GetSelectedItemsFromCheckListBox(popup.PropertiesListBody);
                bool includeAttachments = popup.IncludeAttachmentsCheckBox.IsChecked.Value;

                var pd = new ProgressDialog();

                string uiCulture = Thread.CurrentThread.CurrentUICulture.ToString();

                pd.Execute(ts =>
                        {
                            Thread.CurrentThread.CurrentUICulture = new CultureInfo(uiCulture);
                            pd.UpdateProgress(0, ResourceHelper.GetResourceString("MSG_RETRIEVING_WORK_ITEMS"), true);
                            SetProjectTextWithUrl(project, string.Format("{0}: {1}", ResourceHelper.GetResourceString("MSG_REPORT_TITLE"), queryName));

                            if (query.IsLinkQuery)
                            {
                                WorkItemLinkInfo[] queryResults = query.RunLinkQuery();

                                // Dictionary<Id,HierarchyLevel> 
                                var hierarchy = new Dictionary<int, int>();
                                int hierarchyLevel = 1;
                                int itemsCount = queryResults.Count();
                                for (int i = 0; i < itemsCount; i++)
                                {
                                    try
                                    {
                                        if (!ts.IsCancellationRequested)
                                        {
                                            WorkItemLinkInfo item = queryResults[i];
                                            WorkItem workItem = tfsManager.ItemsStore.GetWorkItem(item.TargetId);
                                            DateTime date = DateTime.MaxValue;

                                            popup.Dispatcher.Invoke((Action)(() =>
                                            {
                                                date = popup.onDate.SelectedDate.GetValueOrDefault(date);
                                            }));

                                            if (date != DateTime.MaxValue)
                                            {
                                                Revision onDateItem = GetLastRevision(workItem, date);

                                                if (onDateItem != null)
                                                {
                                                    pd.UpdateProgress(Convert.ToInt32(i / (decimal)itemsCount * 100), string.Format(ResourceHelper.GetResourceString("REPORT_PROGRESS"), i + 1, itemsCount, onDateItem.WorkItem.Title), false);
                                                    AddHierarchy(hierarchy, ref hierarchyLevel, ref item);

                                                    // TODO:Move to separate Method
                                                    WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, onDateItem, hierarchyLevel);
                                                    //WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, workItem, hierarchyLevel);
                                                }
                                            }
                                            else
                                            {
                                                pd.UpdateProgress(Convert.ToInt32(i / (decimal)itemsCount * 100), string.Format(ResourceHelper.GetResourceString("REPORT_PROGRESS"), i + 1, itemsCount, workItem.Title), false);
                                                AddHierarchy(hierarchy, ref hierarchyLevel, ref item);
                                                WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, workItem, hierarchyLevel);
                                            }
                                        }
                                    }
                                    //catch (DeniedOrNotExistException)
                                    //{
                                    //}
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                                    }
                                }
                            }
                            else
                            {
                                WorkItemCollection queryResults = query.RunQuery();
                                int counter = 0;
                                foreach (WorkItem item in queryResults)
                                {
                                    DateTime date = DateTime.MaxValue;
                                    popup.Dispatcher.Invoke((Action)(() =>
                                    {
                                        date = popup.onDate.SelectedDate.GetValueOrDefault(date);
                                    }));

                                    if (date != DateTime.MaxValue)
                                    {
                                        Revision onDateItem = GetLastRevision(item, date);
                                        if (onDateItem != null)
                                        {
                                            if (!ts.IsCancellationRequested)
                                            {
                                                counter++;

                                                pd.UpdateProgress(Convert.ToInt32(counter / (decimal)queryResults.Count * 100), string.Format(ResourceHelper.GetResourceString("REPORT_PROGRESS"), counter, queryResults.Count, onDateItem.WorkItem.Title), false);
                                                
                                                 //TODO:Move to separate Method
                                                WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, onDateItem);
                                                //WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, item);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (!ts.IsCancellationRequested)
                                        {
                                            counter++;

                                            pd.UpdateProgress(Convert.ToInt32(counter / (decimal)queryResults.Count * 100), string.Format(ResourceHelper.GetResourceString("REPORT_PROGRESS"), counter, queryResults.Count, item.Title), false);

                                            WriteReport(descriptionText, metaDataSelectedItems, bodySelectedItems, includeAttachments, item);
                                            // TODO:Move to separate Method                                           
                                        }
                                    }
                                }
                            }
                        });
                pd.Create(string.Format("{0} {1}", project, ResourceHelper.GetResourceString("MSG_REPORT_TITLE")), Icons.Report);
            }
        }

        private static void AddHierarchy(Dictionary<int, int> hierarchy, ref int hierarchyLevel, ref WorkItemLinkInfo item)
        {
            if (item.SourceId != 0 && !hierarchy.ContainsKey(item.TargetId))
            {
                if (hierarchy.ContainsKey(item.SourceId))
                {
                    hierarchyLevel = hierarchy[item.SourceId] + 1;
                    hierarchy.Add(item.TargetId, hierarchyLevel);
                }
                else
                {
                    hierarchyLevel = 2;
                    hierarchy.Add(item.TargetId, hierarchyLevel);
                }
            }
            else
            {
                hierarchyLevel = 1;
            }
        }

        private void WriteReport(Action<string> descriptionText, List<string> metaDataSelectedItems, List<string> bodySelectedItems, bool includeAttachments, Revision onDateItem, int hierarchyLevel = -1)
        {
            WdBuiltinStyle style = hierarchyLevel != -1 ? ParseEnum<WdBuiltinStyle>(string.Format("wdStyleHeading{0}", hierarchyLevel)) : WdBuiltinStyle.wdStyleHeading1;
            OfficeHelper.SetText(ref document, string.Format("{0} ", onDateItem.WorkItem.Title), style);
            WriteReportMetadata(metaDataSelectedItems, onDateItem.Fields);
            WriteReportBodyData(bodySelectedItems, onDateItem.Fields, descriptionText);
            if (includeAttachments)
            {
                InsertAttachment(onDateItem.Attachments);
            }
        }

        private void WriteReport(Action<string> descriptionText, List<string> metaDataSelectedItems, List<string> bodySelectedItems, bool includeAttachments, WorkItem onDateItem, int hierarchyLevel = -1)
        {
            WdBuiltinStyle style = hierarchyLevel != -1 ? ParseEnum<WdBuiltinStyle>(string.Format("wdStyleHeading{0}", hierarchyLevel)) : WdBuiltinStyle.wdStyleHeading1;
            OfficeHelper.SetText(ref document, string.Format("{0} ", onDateItem.Title), style);
            WriteReportMetadata(metaDataSelectedItems, onDateItem.Fields);
            WriteReportBodyData(bodySelectedItems, onDateItem.Fields, descriptionText);
            if (includeAttachments)
            {
                InsertAttachment(onDateItem.Attachments);
            }
        }

        private Revision GetLastRevision(WorkItem item, DateTime date)
        {
            var revisionsOnDateList = new List<Revision>();
            date = CorrectDate(date);
            foreach (Revision revision in item.Revisions)
            {
                var changedDate = revision.Fields[CoreField.ChangedDate].Value;
                if ((DateTime)changedDate <= date)
                {
                    revisionsOnDateList.Add(revision);
                }
            }
            if (revisionsOnDateList.Count > 0)
            {
                return GetLatestRevision(revisionsOnDateList);
            }
            return null;
        }

        private DateTime CorrectDate(DateTime date)
        {
            date = date.AddHours(23);
            return date;
        }

        private Revision GetLatestRevision(List<Revision> revisionsOnDateList)
        {
            var date = revisionsOnDateList.Select(x => (x.Fields[CoreField.ChangedDate].Value as DateTime?)).Max();
            return revisionsOnDateList.First(x => (x.Fields[CoreField.ChangedDate].Value as DateTime?) == date);
        }

        public static T ParseEnum<T>(string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }

        #endregion

        #region Methods

        private List<string> GetSelectedItemsFromCheckListBox(ListBox list)
        {
            List<string> selectedItems = list.Items.Cast<CheckBox>().Where(f => f.IsChecked ?? false).Select(f => f.Content).Cast<string>().ToList();
            return selectedItems;
        }

        private void InsertAttachment(AttachmentCollection attachments)
        {
            if (attachments.Count != 0)
            {
                //SetMetadataText("\tAttachments:");
                SetMetadataText("\t" + ResourceHelper.GetResourceString("WORD_REPORT_ATTACHMENTS"));
                var webClient = new TfsWebClient(tfsManager.collection);

                string dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string zipFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".zip");

                // AttachmentsParagraph = null;

                try
                {
                    string filePath = string.Empty;
                    foreach (Attachment attachment in attachments)
                    {
                        string fileName = attachment.Name;

                        filePath = Path.Combine(dir, attachment.Name);
                        //process attachments with same names
                        var counter = 0;
                        while (File.Exists(filePath) && counter < 100) //up to 100 items
                        {
                            counter++;

                            if (counter == 1)
                            {
                                filePath = filePath.Insert(filePath.LastIndexOf(".", StringComparison.Ordinal), string.Format("({0})", counter));
                            }
                            else
                            {
                                filePath = filePath.Replace(string.Format("({0}).", counter - 1), string.Format("({0}).", counter));
                            }
                        }

                        try
                        {
                            webClient.DownloadFile(attachment.Uri, filePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    OfficeHelper.CompressDirectory(dir, zipFileName);
                    InsertFile(zipFileName, "Attachments.zip");
                    File.Delete(zipFileName);
                    Directory.Delete(dir, true);
                }
            }
        }

       
        private void InsertFile(string path, string originalFileName)
        {
            object missing = Type.Missing;
            object pathObj = path;
            object classType = "pngfile";
            object iconLabel = originalFileName;
            object iconFilePath = null;

            var fi = new FileInfo(path);
            classType = OfficeHelper.GetFileType(fi, false);

            Paragraph attachmentsParagraph = document.Paragraphs.Add();
            attachmentsParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            document.Content.InsertParagraphAfter();

            object oTrue = -1; //MsoTriState.msoTrue;
            object oFalse = 0; //MsoTriState.msoFalse;
            object iconIndexObj = missing;

            try
            {

                string iconPath;
                int iconIndex;
                OfficeHelper.ExtractIcon(Path.GetExtension(path), out iconPath, out iconIndex);
                iconFilePath = Environment.ExpandEnvironmentVariables(iconPath);
                iconIndexObj = iconIndex;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (iconFilePath == null)
            {
                try
                {
                    Icon objectIcon = Icon.ExtractAssociatedIcon(path);
                    iconFilePath = String.Format("{0}.ico", path.Replace(".", ""));

                    using (var iconStream = new FileStream((string)iconFilePath, FileMode.Create))
                    {
                        objectIcon.Save(iconStream);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            try
            {
                attachmentsParagraph.Range.InlineShapes.AddOLEObject(ref classType, ref pathObj, ref oFalse, ref oTrue, ref iconFilePath, ref iconIndexObj, ref iconLabel, ref missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SetMetadataText(string text)
        {
            Range par = OfficeHelper.CreateParagraphRange(ref document);
            par.Font.Name = "Calibri";
            par.Font.Size = 11;
            par.Font.Color = WdColor.wdColorBlueGray;
            par.Font.Italic = -1;
            par.Text = text;
            par.ParagraphFormat.SpaceAfter = 0;
            par.ParagraphFormat.SpaceBefore = 0;
            document.Content.InsertParagraphAfter();
        }


        private void WriteReportBodyData(IEnumerable<string> selectedProperties, FieldCollection itemFields, Action<string> descriptionText)
        {
            foreach (string prop in selectedProperties)
            {
                var fi = itemFields.Cast<Field>().FirstOrDefault(f => f.Name == prop);
                if (fi != null)
                {
                    SetMetadataText(string.Format("\t{0}: ", fi.Name));
                    try
                    {
                        var val = itemFields[fi.Name].Value.ToString();
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            if (TfsManager.Instance.tfsVersion == TfsVersion.Tfs2010)
                            {
                                var doc = GetNormalizedImages(itemFields[fi.Name].Value.ToString());
                                descriptionText(doc.TrimEnd());
                            }
                            else
                            {
                                //ResourceHelper.GetResourceString("WORD_REPORT_ATTACHMENTS")
                                var doc = GetNormalizedDocument(itemFields[fi.Name].Value.ToString());
                                if (prop == "Шаги" || prop == "Steps")
                                {
                                    doc = doc.Replace("&lt;", "<");
                                    doc = doc.Replace("&gt;", ">");
                                    doc = doc.Replace("&amp;", "");
                                    doc = doc.Replace("nbsp;", "");
                                    doc = doc.Replace("lt;", "&lt;");
                                    doc = doc.Replace("gt;", "&gt;");
                                    doc = doc.Replace("amp;", "&amp;");
                                    doc = doc.Replace("<p>", "");
                                    doc = doc.Replace("</p>", "");
                                    doc = doc.Replace("<P>", "");
                                    doc = doc.Replace("</P>", "");
                                    doc = doc.Replace("<DIV>", "");
                                    doc = doc.Replace("</DIV>", "");
                                    doc = doc.Replace("<div>", "");
                                    doc = doc.Replace("</div>", "");
                                    //doc = doc.Replace("</PARAMETERIZEDSTRING><PARAMETERIZEDSTRING", "</PARAMETERIZEDSTRING> | <PARAMETERIZEDSTRING");
                                    //doc = doc.Replace("</STEP><STEP ", "</STEP><br><STEP ");

                                    doc = doc.Replace("<STEP id=1 ", "<table border='0' width='100%'><tr><td width='50%'><i>" + ResourceHelper.GetResourceString("WORD_PEPORT_ACTION") + "</i></td><td width='50%'><i>"
                                        + ResourceHelper.GetResourceString("WORD_PEPORT_EXPECTED_RESULT") + "</i></td></tr><tr><td width='50%'><STEP id=1 ");
                                    doc = doc.Replace("</PARAMETERIZEDSTRING><PARAMETERIZEDSTRING", "</PARAMETERIZEDSTRING></td><td><PARAMETERIZEDSTRING");
                                    doc = doc.Replace("</STEP><STEP ", "</STEP></td></tr><tr><td><STEP ");
                                    doc = doc.Replace("</STEP></STEPS>", "</STEP></STEPS></td></tr></table>"); 
                                }
                                if (!string.IsNullOrWhiteSpace(doc))
                                {
                                    descriptionText(doc.TrimEnd());
                                }
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
        public void WriteReportBodyDataDescription(int id, Action<string, Comment> descriptionText, Comment comment)
        {
         
             WorkItem wi = tfsManager.GetWorkItem(id);
                var fi = wi.Fields["Description"];
                if (fi != null)
                {
                    //SetMetadataText(string.Format("\t{0}: ", fi.Name));
                    try
                    {
                        var val = wi.Fields["Description"].Value.ToString();
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            if (TfsManager.Instance.tfsVersion == TfsVersion.Tfs2010)
                            {
                                var doc = GetNormalizedImages(wi.Fields["Description"].Value.ToString());
                                descriptionText(doc.TrimEnd(), comment);
                            }
                            else
                            {
                                //ResourceHelper.GetResourceString("WORD_REPORT_ATTACHMENTS")
                                var doc = GetNormalizedDocument(wi.Fields["Description"].Value.ToString());
                                
                                if (!string.IsNullOrWhiteSpace(doc))
                                {
                                    descriptionText(doc.TrimEnd(), comment);
                                }
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            
        }

        public string GetNormalizedImages(string text)
        {
            string removeString = "\n";
            if (text.Contains("See attachments:"))
            {

                int index = text.IndexOf(removeString);
                string temp = index < 0 ? text : text.Remove(index, removeString.Length);
                temp = Regex.Replace(temp, "^.*?Name", "Name");
                temp = temp.TrimEnd(')');
                string[] firstArray = temp.Split(',');
                List<string[]> myList = new List<string[]>();
                text += "<p>";
                string collectionName = TfsManager.Instance.collection.ToString();
                for (int i = 0; i < firstArray.Length; i++)
                {
                    string[] finalArray = firstArray[i].Split(' ');
                    for (int j = 0; j < finalArray.Length; j++)
                        finalArray[j] = finalArray[j].Split('=').LastOrDefault();
                    myList.Add(finalArray);
                    AddImageToDescription(finalArray, collectionName, ref text);
                }
            }
            int startIndex = text.IndexOf("(See attachments");
            if (startIndex != -1)
            {
                int endIndex = text.IndexOf(')');
                text = text.Remove(startIndex, endIndex - startIndex + 1);
                text += "<p>";
            }
            return text;
        }

        private void WriteReportMetadata(List<string> selectedProperties, FieldCollection fields)
        {
            var builder = new StringBuilder();

            if (selectedProperties.Contains("ID"))
            {
                builder.AppendFormat("\t ID: {0} \n", fields[CoreField.Id].Value);
            }

            // Generate Selected properties
            foreach (string property in selectedProperties)
            {
                // do not write ID and Title
                if (property != "ID" && property != "Title")
                {
                    IEnumerable<Field> itemFields = fields.Cast<Field>().Where(fi => fi.Name == property);
                    foreach (Field fi in itemFields)
                    {
                        if (fi.Name == "Tags")
                        {
                            builder.AppendFormat("\t{0}: {1} \n", fi.Name, ((string)fi.Value).Replace(";", "; "));
                        }
                        else
                        {
                            builder.AppendFormat("\t{0}: {1} \n", fi.Name, fi.Value);
                        }
                    }
                }
            }
            if (builder.Length == 0) builder.Length = 2;
            SetMetadataText(builder.Replace("\n", "", builder.Length - 2, 2).ToString());
        }

        private void AddImageToDescription(string[] finalArray,string collectionName,ref string text)
        {
            if (finalArray[0].Contains(".jpg"))
            {
                text +=
                    string.Format(
                        "<img width=643 height=362 src=\"{0}/WorkItemTracking/v1.0/AttachFileHandler.ashx?FileID={1}&FileName={2}\">",
                        collectionName,
                        finalArray[1], finalArray[0]);
            }
            if (finalArray[0].Contains(".png"))
            {
                text +=
                    string.Format(
                        "<img width=300 height=200 src=\"{0}/WorkItemTracking/v1.0/AttachFileHandler.ashx?FileID={1}&FileName={2}\">",
                        collectionName,
                        finalArray[1], finalArray[0]);
            }
        }

        

        private void SetProjectTextWithUrl(string projectName, string title)
        {
            Object styleTitle = WdBuiltinStyle.wdStyleTitle;
            var pt = OfficeHelper.CreateParagraphRange(ref document);
            pt.Text = projectName + " " + title;
            pt.set_Style(ref styleTitle);
            pt.SetRange(pt.Characters.First.Start, pt.Characters.First.Start + projectName.Length);
            pt.Select();

            var link = (tfsManager.tfsVersion == TfsVersion.Tfs2011 ? string.Format("{0}/{1}", TfsManager.Instance.collection.Uri, projectName) : string.Format("{0}/web/", TfsManager.Instance.GetTfsUrl()));
            pt.Hyperlinks.Add(pt, link);
            document.Content.InsertParagraphAfter();

            var RevisionPar = OfficeHelper.CreateParagraphRange(ref document);
            if (revisionDate != null)
            {
                RevisionPar.Text = string.Format("{0}: {1}",WordToTFS.Properties.Resources.QueryReport_RevisionDate_Text,revisionDate.Value.ToShortDateString());
                RevisionPar.Font.Name = "Calibri";
                RevisionPar.Font.Size = 11;
                RevisionPar.Font.Color = WdColor.wdColorBlueGray;
                RevisionPar.Font.Italic = -1;
                document.Content.InsertParagraphAfter();
            }
           

           

        }

        /// <summary>
        /// The get normalized document.
        /// </summary>
        /// <param name="text">
        /// The text.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public string GetNormalizedDocument(string text)
        {
            string dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            Path.Combine(dir, Guid.NewGuid().ToString("N"));

            IHTMLDocument2 doc = getHTMLDocument(text.SafeHtml());

            var imagesToDelete = new List<IHTMLImgElement>();

            foreach (IHTMLImgElement image in doc.images)
            {
                try
                {
                    if (image.src.TrimStart(' ').StartsWith("data:"))
                    {
                        var regex = new Regex(@"data:(?<mime>[\w/]+);(?<encoding>\w+),(?<data>.*)", RegexOptions.Compiled);
                        var match = regex.Match(image.src);
                        var mime = match.Groups["mime"].Value;
                        var encoding = match.Groups["encoding"].Value;
                        var data = match.Groups["data"].Value;

                        var mimeSpl = mime.Split('/');

                        if (mimeSpl.Length > 1)
                        {
                            mime = mimeSpl[mimeSpl.Length - 1];
                        }

                        var imgName = Guid.NewGuid().ToString("N") + "." + mime;
                        var imgPath = Path.Combine(dir, imgName);

                        File.WriteAllBytes(imgPath, Convert.FromBase64String(data));
                        image.src = string.Format("file:///{0}", imgPath.Replace('\\', '/'));
                    }
                    else if (string.IsNullOrEmpty(image.src) || string.IsNullOrWhiteSpace(image.src))
                    {
                        imagesToDelete.Add(image);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            foreach (IHTMLDOMNode img in imagesToDelete)
            {
                try
                {
                    img.parentNode.removeChild(img);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            return doc.body.innerHTML;
        }


        /// <summary>
        /// The get html document.
        /// </summary>
        /// <param name="html">
        /// The html.
        /// </param>
        /// <returns>
        /// The <see cref="IHTMLDocument2"/>.
        /// </returns>
        private IHTMLDocument2 getHTMLDocument(string html)
        {
            var doc = (IHTMLDocument2)new HTMLDocument();
            doc.write(new object[] { html });
            doc.close();
            return doc;
        }



        #endregion
    }
}
