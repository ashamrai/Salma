using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Xml;
using WordToTFS.View;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.Office.Interop.Word;
using WordToTFS.Properties;
using mshtml;
using DataFormats = System.Windows.DataFormats;
using Field = Microsoft.TeamFoundation.WorkItemTracking.Client.Field;
using IDataObject = System.Windows.Forms.IDataObject;
using MessageBox = System.Windows.MessageBox;
using Revision = Microsoft.TeamFoundation.WorkItemTracking.Client.Revision;

//ResourceHelper.GetResourceString("ALL")
namespace WordToTFS
{
    using Microsoft.VisualStudio.Services.Common;
    using System.Diagnostics;
    using System.Windows.Controls;

    /// <summary>
    ///  The tfs manager.
    /// </summary>
    public class TfsManager
    {
        #region Fields

        /// <summary>
        /// The instance.
        /// </summary>
        private static volatile TfsManager instance = null;

        private static object syncRoot = new Object();

        /// <summary>
        /// The collection.
        /// </summary>
        public TfsTeamProjectCollection collection;

        /// <summary>
        ///     The connect uri.
        /// </summary>
        private Uri connectUri;

        /// <summary>
        ///     The work item store.
        /// </summary>
        private WorkItemStore itemsStore;

        /// <summary>
        ///     The nodes.
        /// </summary>
        private List<CatalogNode> nodes;

        #endregion

        #region Public Properties


        /// <summary>
        ///     Gets or sets the work item store.
        /// </summary>
        //public WorkItemStore ItemsStore
        //{
        //    get { return itemsStore ?? (itemsStore = collection.GetService<WorkItemStore>()); }

        //    set { itemsStore = value; }
        //}

        public string StateName
        {
            get;

            set;
        }

        public WorkItemStore ItemsStore
        {
            get;

            set;
        }

        /// <summary>
        ///     Gets or sets the TFS version.
        /// </summary>
        public TfsVersion tfsVersion { get; set; }

        #endregion

        private TfsManager()
        {
        }

        public ICredentials Credential { set; get; }

        public static TfsManager Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null) instance = new TfsManager();
                    }
                }

                return instance;
            }
        }

        #region Public Methods and Operators

        byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        string GetString(byte[] bytes)
        {
            char[] chars = new char[bytes.Length / sizeof(char)];
            System.Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length);
            return new string(chars);
        }

        /// <summary>
        /// The add details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="dataObject">
        /// The data Object.
        /// </param>
        public void AddDetailsForWorkItem(int id, string fieldName, IDataObject dataObject, out bool comment)
        {
            //var html = dataObject.GetData(DataFormats.Html).ToString();
            string html = ClipboardHelper.Instanse.CopyFromClipboard();

            WorkItem workItem = ItemsStore.GetWorkItem(id);
            string attachmentName = string.Empty;
            html = html.CleanHeader();
            IDictionary<string, string> links = html.GetLinks();

            foreach (var link in links)
            {
                if (File.Exists(link.Value) &&
                    !(new FileInfo(link.Value).Extension == ".thmx" || new FileInfo(link.Value).Extension == ".xml" ||
                      new FileInfo(link.Value).Extension == ".mso"))
                {
                    var attach = new Attachment(link.Value);
                    if (tfsVersion == TfsVersion.Tfs2010)
                    {

                        var attachIndex = workItem.Attachments.Add(attach);
                        workItem.Save();
                        attachmentName += string.Format("Name={0} Id={1},", workItem.Attachments[attachIndex].Name, workItem.Attachments[attachIndex].Id);
                        string uri = attach.Uri.ToString();
                        html = html.Replace(link.Key, uri);
                    }
                    else
                    {
                        workItem.Attachments.Add(attach);
                        workItem.Save();
                        string uri = attach.Uri.ToString();
                        html = html.Replace(link.Key, uri);
                    }

                }
                else
                {
                    html = html.Replace(link.Key, string.Empty);
                }
            }

            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            //Field fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Id == 10265);
            Field fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == "Steps");
            var fields = workItem.Fields;
            if (fieldSteps == null)
            {
                fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == "Шаги");
                //fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Id == 10032);
            }
            comment = false;
            if (itemField != null)
            {
                if (fieldSteps != null && fieldName == fieldSteps.Name)
                {
                        AddStep(dataObject, workItem, out comment, true);
                }
                else
                {

                    if (tfsVersion == TfsVersion.Tfs2010)
                    {
                        if (itemField.FieldDefinition.FieldType == FieldType.Html)
                        {
                            itemField.Value += html.ClearComments();
                        }
                        else
                        {
                            string temp = Regex.Replace(dataObject.GetData(DataFormats.Text).ToString(), @"\s+", " ");
                            itemField.Value += "\r" + temp;
                            if (attachmentName != string.Empty)
                                itemField.Value += string.Format("\r({0}: {1})", Resources.SeeAttachments_Text, attachmentName.TrimEnd(','));
                        }

                    }
                    else
                    {
                        itemField.Value += html.ClearComments();
                    }
                    workItem.Save();
                    comment = true;
                }


            }
        }

        /// <summary>
        /// The add step for work item "Test Case".
        /// </summary>
        /// <param name="dataObject"></param>
        /// <param name="workItem"></param>
        public void AddStep(IDataObject dataObject, WorkItem workItem, out bool comment,bool isAddStep)
        {
            var popup = new StepActionsResult();
            string temp = Regex.Replace(dataObject.GetData(DataFormats.Text).ToString(), @"\s+", " ");
            popup.action.Text = temp;
            popup.Create(null, Icons.AddDetails);
            ITestManagementService testService = collection.GetService<ITestManagementService>();
            var project = testService.GetTeamProject(workItem.Project.Name);
            var testCase = project.TestCases.Find(workItem.Id);
            var step = testCase.CreateTestStep();

            if (!popup.IsCanceled)
            {
                switch (tfsVersion)
                {
                    case TfsVersion.Tfs2011:

                        step.Title = "<div><p><span>" + popup.action.Text + "</span></p></div>";
                        step.ExpectedResult = "<div><p><span>" + popup.expectedResult.Text + "</span></p></div>";


                        //step.Title = popup.action.Text;
                        //step.ExpectedResult = popup.expectedResult.Text;

                        break;
                    case TfsVersion.Tfs2010:

                        step.Title = popup.action.Text;
                        step.ExpectedResult = popup.expectedResult.Text;

                        break;
                }
                if (isAddStep)
                {
                    testCase.Actions.Add(step);
                }
                else
                {
                    testCase.Actions.Clear();
                    testCase.Actions.Add(step);
                }
                testCase.Save();
                workItem.Save();
                comment = true;
            }
            else
            {
                comment = false;
            }
        }

        /// <summary>
        /// The add details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="bitmapSource">
        /// The bitmap source.
        /// </param>
        public void AddDetailsForWorkItem(int id, string fieldName, BitmapSource bitmapSource)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);

            string temp = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) + @"\image.png";
            using (var ms = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapSource));
                enc.Save(ms);
                File.WriteAllBytes(temp, ms.ToArray());
            }

            var attach = new Attachment(temp);
            workItem.Attachments.Add(attach);
            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            if (itemField != null)
            {
                itemField.Value += string.Format(@"<img src='{0}' />", attach.Uri);
            }

            workItem.Save();
        }

        /// <summary>
        /// The add details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="ms">
        /// The memory stream.
        /// </param>
        public void AddDetailsForWorkItem(int id, string fieldName, MemoryStream ms)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);

            string temp = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) + @"\image.png";
            File.WriteAllBytes(temp, ms.ToArray());

            var attach = new Attachment(temp);
            var attachIndex = workItem.Attachments.Add(attach);
            workItem.Save();
            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            if (itemField != null)
            {
                switch (tfsVersion)
                {
                    case TfsVersion.Tfs2011:
                        itemField.Value += string.Format(@"<img src='{0}' />", workItem.Attachments[attachIndex].Uri);
                        break;
                    case TfsVersion.Tfs2010:
                        itemField.Value += string.Format("\r({0} : {1})\r",ResourceHelper.GetResourceString("SEE_ATTACHEMENTS"), workItem.Attachments[attachIndex].Name);
                        break;
                }
            }

            workItem.Save();
        }

        /// <summary>
        /// The add history to work item.
        /// </summary>
        /// <param name="workItemId">
        /// The work item id.
        /// </param>
        /// <param name="historyText">
        /// The history text.
        /// </param>
        public void AddHistoryToWorkItem(int workItemId, string historyText)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(workItemId);
            workItem.History = historyText;
            workItem.Save();
        }


        /// <summary>
        /// The change collection.
        /// </summary>
        /// <param name="collectionName">
        /// The collection name.
        /// </param>
        public void ChangeCollection (string collectionName)
        {
            bool _isvsservice = connectUri.ToString().Contains("visualstudio.com");
            collection = new TfsTeamProjectCollection(new Uri(string.Format(@"{0}/{1}", connectUri, (_isvsservice) ? "" : collectionName)), Credential);
            collection.EnsureAuthenticated();
            ItemsStore = (WorkItemStore)collection.GetService(typeof(WorkItemStore));
        }

        /// <summary>
        /// The connect.
        /// </summary>
        /// <param name="uri">
        /// The uri.
        /// </param>
        /// <param name="credentials">
        /// The credentials.
        /// </param>
        /// <returns>
        /// is connected
        /// </returns>
        public bool Connect(Uri uri, ICredentials credentials, bool anotherUser)
        {

            List<CatalogNode> catalogNodes = GetCatalogNodes(uri, credentials, anotherUser);
            
            if (catalogNodes != null)
            {
                nodes = new List<CatalogNode>();
                foreach (CatalogNode node in catalogNodes)
                {
                    try
                    {
                        bool _isvsservice = uri.ToString().Contains("visualstudio.com");
                        using (TfsTeamProjectCollection projCollection = new TfsTeamProjectCollection(new Uri(string.Format(@"{0}/{1}", uri, (_isvsservice)? "" : node.Resource.DisplayName)), Credential))
                        {
                            projCollection.EnsureAuthenticated();
                            projCollection.Connect(ConnectOptions.None);
                            nodes.Add(node);
                        }                        
                    }
                    catch (TeamFoundationServiceUnavailableException ex)
                    {
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                    }
                }

                if (nodes.Count > 0)
                {
                    ChangeCollection(nodes[0].Resource.DisplayName);
                    tfsVersion = collection.GetVersion();
                    return true;
                }
            }

            return false;
        }

        public void Disconnect()
        {
            if (collection != null)
            {
                collection.Disconnect();
                collection.Dispose();
            }
        }

        /// <summary>
        /// The execute query.
        /// </summary>
        /// <param name="query">
        /// The query.
        /// </param>
        /// <returns>
        /// The <see cref="WorkItemCollection"/>.
        /// </returns>
        public WorkItemCollection ExecuteQuery(string query)
        {
            return ItemsStore.Query(query);
        }

        /// <summary>
        /// The execute tree query.
        /// </summary>
        /// <param name="query">
        /// The query.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<WorkItemLinkInfo> ExecuteTreeQuery(string query)
        {
            var treeQuery = new Query(ItemsStore, query);
            return treeQuery.RunLinkQuery().ToList();
        }

        /// <summary>
        /// The generate report.
        /// </summary>
        /// <param name="project">
        /// The project.
        /// </param>
        /// <param name="reportTitle">
        /// The report title.
        /// </param>
        /// <param name="workItemTitle">
        /// The work item title.
        /// </param>
        /// <param name="itemDetails">
        /// The item details.
        /// </param>
        /// <param name="descriptionText">
        /// The description text.
        /// </param>
        /// <param name="insertFile">
        /// The insert file.
        /// </param>


        /// <summary>
        ///     The get all work item links types.
        /// </summary>
        /// <returns>
        ///     The <see cref="WorkItemLinkTypeCollection" />.
        /// </returns>
        public WorkItemLinkTypeCollection GetAllWorkItemLinksTypes()
        {
            return ItemsStore.WorkItemLinkTypes;
        }

        /// <summary>
        /// The get all work items for project.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<WorkItem> GetAllWorkItemsForProject(string projectName)
        {
            string wiql = string.Format(@"SELECT * FROM WorkItems WHERE [System.TeamProject] = '{0}'", projectName);
            return ExecuteQuery(wiql).Cast<WorkItem>().ToList();
        }

        /// <summary>
        /// The get html fields by item id.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<string> GetHtmlFieldsByItemId(int id)
        {
            try
            {
                WorkItem workItem = ItemsStore.GetWorkItem(id);
                switch (tfsVersion)
                {
                    case TfsVersion.Tfs2011:
                        return
                            workItem.Fields.Cast<Field>()
                                    .Where(f => f.FieldDefinition.FieldType == FieldType.Html)
                                    .Select(f => f.Name)
                                    .ToList();
                    case TfsVersion.Tfs2010:
                        {
                            return
                                workItem.Fields.Cast<Field>()
                                        .Where(
                                            f =>
                                            f.FieldDefinition.FieldType == FieldType.Html &&
                                            f.FieldDefinition.Name != SalmaConstants.TFS.LOCAL_DATA_SOURCE ||
                                            (f.FieldDefinition.FieldType == FieldType.PlainText &&
                                             f.FieldDefinition.Name !=
                                             SalmaConstants.TFS.PROJECT_SERVER_SYNC_ASSIGMENT_DATA) &&
                                            f.FieldDefinition.Name != SalmaConstants.TFS.LOCAL_DATA_SOURCE)
                                        .Select(f => f.Name)
                                        .ToList();
                        }
                }

                return null;
            }
            catch (DeniedOrNotExistException e)
            {
                return null;
            }
            catch (ArgumentOutOfRangeException e)
            {
                return null;
            }
        }

        /// <summary>
        /// Get workitem details default field
        /// </summary>
        /// <param name="id">
        /// workitem id.
        /// </param>
        /// <returns>
        /// field Name
        /// </returns>
        public string GetDefaultDetailsFieldName(int id)
        {
            try
            {
                if (id > 0)
                {
                    string fieldType;
                    var workItem = ItemsStore.GetWorkItem(id);
                    using (var stringReader = new StringReader(workItem.DisplayForm))
                    {
                        using (var reader = XmlReader.Create(stringReader))
                        {
                            reader.ReadToFollowing("TabGroup");
                            reader.ReadToFollowing("Control");
                            fieldType = reader.GetAttribute("FieldName");
                        }
                    }
                    return
                        workItem.Fields.Cast<Field>()
                                .Where(f => f.FieldDefinition.ReferenceName == fieldType)
                                .Select(f => f.Name)
                                .FirstOrDefault() ?? string.Empty;
                }
                return string.Empty;
            }
            catch (DeniedOrNotExistException e)
            {
                return null;
            }
        }

        /// <summary>
        /// The get project.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <returns>
        /// The <see cref="Project"/>.
        /// </returns>
        public Project GetProject(string projectName)
        {
            return ItemsStore.Projects[projectName];
            
        }

        /// <summary>
        /// The get projects for current tpc.
        /// </summary>
        /// <param name="index">
        /// The index.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<string> GetProjectsForCurrentTPC(int index)
        {
            return nodes[index].QueryChildren(new[] { CatalogResourceTypes.TeamProject }, false, CatalogQueryOptions.None).Select(f => f.Resource.DisplayName).ToList();
        }

        /// <summary>
        /// Get areas for project
        /// </summary>
        /// <param name="pProjectName"></param>
        /// <returns></returns>
        public List<string> GetAreasForProject(string pProjectName)
        {
            List<string> resultList = new List<string>();

            if (ItemsStore == null) return null;

            resultList.Add(ItemsStore.Projects[pProjectName].Name);

            foreach (Node area in ItemsStore.Projects[pProjectName].AreaRootNodes)
            {
                resultList.Add(area.Path);
                if (area.ChildNodes.Count > 0) GetChildAreaNodes(resultList, area.ChildNodes);
            }

            return resultList;
        }

        /// <summary>
        /// Get child areas
        /// </summary>
        /// <param name="resultList"></param>
        /// <param name="nodeCollection"></param>
        private void GetChildAreaNodes(List<string> resultList, NodeCollection nodeCollection)
        {
            foreach (Node area in nodeCollection)
            {
                resultList.Add(area.Path);
                if (area.ChildNodes.Count > 0) GetChildAreaNodes(resultList, area.ChildNodes);
            }
        }

        public List<string> GetCollectionLinkEnds()
        {
            List<string> resultList = new List<string>();

            resultList.Add(string.Empty);

            foreach (WorkItemLinkType type in ItemsStore.WorkItemLinkTypes)
            {
                resultList.Add(type.ForwardEnd.Name);
                if (!resultList.Contains(type.ReverseEnd.Name)) resultList.Add(type.ReverseEnd.Name);
            }

            return resultList;
        }

        /// <summary>
        /// The get queries.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<QueryItem> GetQueries(string projectName)
        {
            return ItemsStore.Projects[projectName].QueryHierarchy.ToList();
        }

        /// <summary>
        ///     The get tpc.
        /// </summary>
        /// <returns>
        ///     The <see cref="List" />.
        /// </returns>
        public List<string> GetTeamProjectCollection()
        {
            return nodes.Select(f => f.Resource.DisplayName).ToList();
        }

        /// <summary>
        ///     The get tfs url.
        /// </summary>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string GetTfsUrl()
        {
            return connectUri.ToString();
        }

        /// <summary>
        ///     The get user display name.
        /// </summary>
        /// <returns>
        ///     The <see cref="string" />.
        /// </returns>
        public string GetUserDisplayName()
        {
            return ItemsStore.UserDisplayName;
        }

        /// <summary>
        /// The get work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <returns>
        /// The <see cref="WorkItem"/>.
        /// </returns>
        public WorkItem GetWorkItem(int id)
        {
            try
            {
                return ItemsStore.GetWorkItem(id);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public bool GetWorkItemTextInfo(int id, out string Title, out string WIType)
        {
            Title = "";
            WIType = "";

            WorkItem wi = GetWorkItem(id);

            if (wi != null)
            {
                Title = wi.Title;
                WIType = wi.Type.Name;
            }
            else
                return false;

            return true;
        }

        public List<string> GetWorkItemHistory(int workItemId)
        {
            var log = new List<string>();
            WorkItem workItem = ItemsStore.GetWorkItem(workItemId);
            foreach (Revision revision in workItem.Revisions)
            {
                log.AddRange(from Field field in workItem.Fields where revision.Fields[field.Name].Value != null where !string.IsNullOrEmpty(revision.Fields[field.Name].Value.ToString()) select field.Name + ": " + revision.Fields[field.Name].Value);
                log.Add("-------------------------------------");
            }
            return log;

            
          
        }

        public int CompareDate(DateTime wordComment, DateTime revision)
        {

            var comparedDate = revision.Subtract(wordComment).Minutes;
            
            return comparedDate;
        }

        public bool GetWorkItemFieldHistory(int workItemId, string FieldName, DateTime Date) //Check if Text was changed in WI revisions, comparing comment data created with WI Changed Date
        {
            WorkItem workItem = ItemsStore.GetWorkItem(workItemId);
            bool historyIsChanged = false;

           if (FieldName != String.Empty && FieldName != "Title")
           { 
                         

                          
                var AllRevisions = from r in workItem.Revisions.OfType<Revision>()
                                   where ((CompareDate(Date, DateTime.Parse(r.Fields["System.ChangedDate"].Value.ToString())) >= 1) && (r.Fields[FieldName].IsChangedInRevision == true))
                    group r by r.Fields[FieldName].IsChangedInRevision
                    into g

                    select new
                    {
                        IsChangedInRevision = g.Key,
                        Count = g.Count(),
                    };
                int rev =  AllRevisions.Count();
                if (rev > 0)
                {
                    return historyIsChanged = true;
                }
            }

            if (FieldName == "Title")
            {

                foreach (Revision revision in workItem.Revisions)
                {
                    if (revision.Index > 0)
                    {
                        if (revision.Fields["System.Title"].IsChangedInRevision)
                        {
                            DateTime revisionTime =
                                DateTime.Parse(revision.Fields["System.ChangedDate"].Value.ToString());
                            var comparedDate = revisionTime.Subtract(Date);
                            if (comparedDate.Minutes >= 1)
                            {
                                return historyIsChanged = true;
                            }
                        }

                    }
                }

            }
            
            //var AllRevisions = from r in workItem.Revisions.OfType<Revision>()
                //    where (DateTime.Parse(r.Fields["Changed Date"].Value.ToString()) > Date)
                //                   && (r.Fields["Title"].IsChangedInRevision == true)
                //    group r by r.Fields["Title"].IsChangedInRevision
                //    into g

                //    select new
                //    {
                //        IsChangedInRevision = g.Key,
                //        Count = g.Count(),
                //    };

                //int rev = AllRevisions.Count();
                //if (rev > 0)
                //{
                //    return historyIsChanged = true;
                //}
            return historyIsChanged = false;
        }





        /// <summary>
        /// The get work item link.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public string GetWorkItemLink(int id, string projectName)
        {
            switch (tfsVersion)
            {
                case TfsVersion.Tfs2011:
                    string _service_uri = ItemsStore.TeamProjectCollection.Uri.ToString();
                    while (_service_uri.Length > 0 && _service_uri[_service_uri.Length - 1] == '/') _service_uri = _service_uri.Substring(0, _service_uri.Length - 1);
                    //return string.Format(@"{0}/{1}/_workitems#_a=edit&id={2}", ItemsStore.TeamProjectCollection.Uri, projectName, id);
                    return string.Format(@"{0}/{1}/_workitems/edit/{2}", _service_uri, projectName, id);
                case TfsVersion.Tfs2010:
                    return string.Format(@"{0}/web/UI/Pages/WorkItems/WorkItemEdit.aspx?id={1}", ItemsStore.TeamProjectCollection.Uri, id);
            }

            return null;
        }

        /// <summary>
        /// Get Work Item States by Type.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <param name="itemType">
        /// The work item type.
        /// </param>
        /// <returns>
        /// available states list
        /// </returns>
        public List<string> GetWorkItemStatesByType(string projectName, string itemType)
        {
            WorkItemType workItemType = ItemsStore.Projects[projectName].WorkItemTypes[itemType];
            var item = new WorkItem(workItemType);
            //return item.Fields["State"].FieldDefinition.AllowedValues.Cast<string>().ToList();
            StateName = item.Fields.GetById(2).Name;
            return item.Fields.GetById(2).FieldDefinition.AllowedValues.Cast<string>().ToList();
            
        }
        //TODO: Review and delete.
        /// <summary>
        /// The get work item text.
        /// </summary>
        /// <param name="workItemId">
        /// The work item id.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public string GetWorkItemText(int workItemId)
        {
            int linksCount = 0;
            return GetWorkItemText(workItemId, out linksCount);
        }
        //TODO: Move to Add-In
        /// <summary>
        /// The get work item text.
        /// </summary>
        /// <param name="workItemId">
        /// The work item id.
        /// </param>
        /// <param name="linksCount">
        /// The links count.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        /// 
        public string GetWorkItemDescription(int workItemId, string fieldName)
        {
            WorkItem wi = GetWorkItem(workItemId);
            return wi.Fields[fieldName].Value.ToString();
        }

        public string GetWorkItemTitle(int workItemId)
        {
            WorkItem wi = GetWorkItem(workItemId);
            return wi.Title.ToString();
        }

        public string GetControlText(int workItemId, string fieldName)
        {
            WorkItem wi = GetWorkItem(workItemId);
            return string.Format("{0} {1} - {2} - {3}", wi.Type.Name, wi.Id, wi.Fields[fieldName].Name, wi.State);
        }

        public string GetWorkItemText(int workItemId, out int linksCount)
        {
            WorkItem wi = GetWorkItem(workItemId);
            linksCount = wi.Links.Count;
            ResourceHelper.GetResourceString("ERROR_TEXT");
            return string.Format("\n{0}{1}\n{2}{3}\n{4}{5}\n{6}{7}\n{8}{9}", ResourceHelper.GetResourceString("CREATED_BY"), GetUserDisplayName(), ResourceHelper.GetResourceString("TYPE"),
                wi.Type.Name, ResourceHelper.GetResourceString("TITLE"), wi.Title, ResourceHelper.GetResourceString("STATE"), wi.State, ResourceHelper.GetResourceString("PROJECT"), wi.Project.Name);
            //return string.Format("\nFARAAAA: {0}\nType: {1}\nTitle: {2}\nStatus: {3}\nProject: {4}", GetUserDisplayName(), wi.Type.Name, wi.Title, wi.State, wi.Project.Name);
        }

        /// <summary>
        /// The get work items type for current project.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<string> GetWorkItemsTypeForCurrentProject(string projectName)
        {
            var sortedList = (from WorkItemType workItemType in ItemsStore.Projects[projectName].WorkItemTypes where workItemType.Name != "Code Review Request" && workItemType.Name != "Code Review Response" && workItemType.Name != "Feedback Request" && workItemType.Name != "Feedback Response" && workItemType.Name != "Shared Steps" select workItemType.Name).ToList();
            sortedList = (from string workItemType in sortedList where workItemType != "Запрос на проверку кода" && workItemType != "Ответ на проверку кода" && workItemType != "Ответ на отзыв" && workItemType != "Запрос отзыва" && workItemType != "Общие шаги" select workItemType).ToList();
            sortedList.Sort();
            return sortedList;
        }

        /// <summary>
        /// Replace details for work item
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fieldName"></param>
        /// <param name="data"></param>
        public void ReplaceDetailsForWorkItem(int id, string fieldName, string data)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);
           
            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase));

            if (fieldName.Equals("system.title", StringComparison.InvariantCultureIgnoreCase))
                itemField = workItem.Fields["System.Title"];

            if (itemField != null)
            {
                itemField.Value = null;
                itemField.Value += data;
                workItem.Save();
            }
        }

        /// <summary>
        /// Is step
        /// </summary>
        /// <param name="id"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public bool IsStep(int id, string fieldName)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);

            Field fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == "Steps");

            if (fieldSteps == null)
                fieldSteps = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == "Шаги");

            if (fieldSteps != null && fieldName == fieldSteps.Name)
                return true;

            return false;
        }

        /// <summary>
        /// Add step
        /// </summary>
        /// <param name="id"></param>
        /// <param name="text"></param>
        public void AddStep(int id, string text)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);

            var popup = new StepActionsResult();
            popup.action.Text = text;
            popup.Create(null, Icons.AddDetails);

            ITestManagementService testService = collection.GetService<ITestManagementService>();
            var project = testService.GetTeamProject(workItem.Project.Name);
            var testCase = project.TestCases.Find(workItem.Id);
            var step = testCase.CreateTestStep();

            if (!popup.IsCanceled)
            {
                switch (tfsVersion)
                {
                    case TfsVersion.Tfs2011:

                        step.Title = "<div><p><span>" + popup.action.Text + "</span></p></div>";
                        step.ExpectedResult = "<div><p><span>" + popup.expectedResult.Text + "</span></p></div>";

                        break;
                    case TfsVersion.Tfs2010:

                        step.Title = popup.action.Text;
                        step.ExpectedResult = popup.expectedResult.Text;

                        break;
                }

                //testCase.Actions.Clear();
                testCase.Actions.Add(step);
                testCase.Save();
                workItem.Save();
            }
        }

        /// <summary>
        /// The replace details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="dataObject">
        /// The data Object.
        /// </param>
        public void ReplaceDetailsForWorkItem(int id, string fieldName, IDataObject dataObject, out bool comment)
        {
            //var html = dataObject.GetData(DataFormats.Html).ToString();
            string html = ClipboardHelper.Instanse.CopyFromClipboard();

            WorkItem workItem = ItemsStore.GetWorkItem(id);
            workItem.Attachments.Clear();
            html = html.CleanHeader();
            string attachmentName = string.Empty;
            IDictionary<string, string> links = html.GetLinks();
            foreach (var link in links)
            {
                if (File.Exists(link.Value) &&
                    !(new FileInfo(link.Value).Extension == ".thmx" || new FileInfo(link.Value).Extension == ".xml"
                    || new FileInfo(link.Value).Extension == ".mso"))
                {
                    var attach = new Attachment(link.Value);
                    if (tfsVersion == TfsVersion.Tfs2010)
                    {
                        var attachIndex = workItem.Attachments.Add(attach);
                        workItem.Save();
                        attachmentName += string.Format("Name={0} Id={1},", workItem.Attachments[attachIndex].Name, workItem.Attachments[attachIndex].Id);
                        string uri = attach.Uri.ToString();
                        html = html.Replace(link.Key, uri);
                    }
                    else
                    {
                        workItem.Attachments.Add(attach);
                        workItem.Save();
                        string uri = attach.Uri.ToString();
                        html = html.Replace(link.Key, uri);
                    }
                }
                else
                {
                    html = html.Replace(link.Key, string.Empty);
                }
            }

            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            comment = false;
            if (itemField != null)
            {
                itemField.Value = null;
                    if (tfsVersion == TfsVersion.Tfs2010)
                    {
                        if (itemField.FieldDefinition.FieldType == FieldType.Html)
                        {
                            itemField.Value += html.ClearComments();
                        }
                        else
                        {
                            string temp = Regex.Replace(dataObject.GetData(DataFormats.Text).ToString(), @"\s+", " ");
                            itemField.Value = "\r" + temp;
                            if (attachmentName != string.Empty)
                                itemField.Value += string.Format("\r({0}: {1})", Resources.SeeAttachments_Text,
                                                                 attachmentName.TrimEnd(','));
                        }
                       
                    }
                    else
                    {
                        itemField.Value += html.ClearComments();
                    }
                    workItem.Save();
                    comment = true;
            }

        }

        /// <summary>
        /// The replace details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="bitmapSource">
        /// The bitmap source.
        /// </param>
        public void ReplaceDetailsForWorkItem(int id, string fieldName, BitmapSource bitmapSource)
        {
            WorkItem workItem = GetWorkItem(id);

            string temp = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) + @"\image.png";
            using (var ms = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapSource));
                enc.Save(ms);
                File.WriteAllBytes(temp, ms.ToArray());
            }

            var attach = new Attachment(temp);
            workItem.Attachments.Add(attach);
            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            if (itemField != null)
            {
                itemField.Value = string.Format(@"<img src='{0}' />", attach.Uri);
            }

            workItem.Save();
        }

        /// <summary>
        /// The replace details for work item.
        /// </summary>
        /// <param name="id">
        /// The id.
        /// </param>
        /// <param name="fieldName">
        /// The field name.
        /// </param>
        /// <param name="ms">
        /// The ms.
        /// </param>
        public void ReplaceDetailsForWorkItem(int id, string fieldName, MemoryStream ms)
        {
            WorkItem workItem = ItemsStore.GetWorkItem(id);

            string temp = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) + @"\image.png";
            File.WriteAllBytes(temp, ms.ToArray());

            var attach = new Attachment(temp);
            workItem.Attachments.Clear();
            //workItem.Attachments.Add(attach);
            var attachIndex = workItem.Attachments.Add(attach);
            workItem.Save();
            Field itemField = workItem.Fields.Cast<Field>().FirstOrDefault(f => f.Name == fieldName);
            if (itemField != null)
            {
                switch (tfsVersion)
                {
                    case TfsVersion.Tfs2011:
                        itemField.Value = string.Format(@"<img src='{0}' />", workItem.Attachments[attachIndex].Uri);
                        break;
                    case TfsVersion.Tfs2010:
                        itemField.Value = string.Format("\r({0} {1})\r", ResourceHelper.GetResourceString("SEE_ATTACHEMENTS"), workItem.Attachments[attachIndex].Name);
                        break;
                }
            }

            workItem.Save();
        }

        #endregion

        #region Methods

        /// <summary>
        /// The parse id.
        /// </summary>
        /// <param name="text">
        /// The text.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public static int ParseId(string text)
        {
            string[] splText = text.Split('\r');
            string temp = null;

            for (int i = 0; i < splText.Length; i++)
            {
                if (splText[i].IndexOf(ResourceHelper.GetResourceString("WI_ID")) != -1)
                {
                    temp = splText[i].Split(':').LastOrDefault();
                    break;
                }
            }

            return int.Parse(temp.Trim());
        }

        /// <summary>
        /// Get All Available WorkItem Fields For Project
        /// </summary>
        /// <param name="project">
        /// Project Name
        /// </param>
        /// <returns>
        /// Field Definition List
        /// </returns>
        public List<FieldDefinition> GetAllAvailableWorkItemFieldsForProject(string project)
        {
            var allFields = new List<FieldDefinition>();
            IEnumerable<FieldDefinitionCollection> allFieldDefinitionsCollection = ItemsStore.Projects[project].WorkItemTypes.Cast<WorkItemType>().Select(f => f.FieldDefinitions);

            foreach (FieldDefinitionCollection fcollection in allFieldDefinitionsCollection)
            {
                allFields.AddRange(fcollection.Cast<FieldDefinition>().ToList());
            }

            allFields = allFields.Select(f => f.Name).Distinct().Select(name => allFields.First(f => f.Name == name)).OrderBy(f => f.Name).ToList();

            return allFields;
        }

        /// <summary>
        /// The get catalog nodes.
        /// </summary>
        /// <param name="uri">
        /// The uri.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        private List<CatalogNode> GetCatalogNodes(Uri uri, ICredentials cred, bool anotherUser)
        {
            List<CatalogNode> catalogNodes = null;
            try
            {

                connectUri = uri;
                TfsClientCredentials tfsClientCredentials;
                if (anotherUser)
                {
                    
                    if (connectUri.ToString().Contains(".visualstudio.com"))
                    {
                        tfsClientCredentials = new TfsClientCredentials(false);
                    }
                    else
                    {
                        Microsoft.TeamFoundation.Client.WindowsCredential wcred = new Microsoft.TeamFoundation.Client.WindowsCredential(false);
                        tfsClientCredentials = new TfsClientCredentials(wcred);
                    }

                }
                else
                {
                    if (connectUri.ToString().Contains(".visualstudio.com"))
                    {
                        tfsClientCredentials = new TfsClientCredentials();                        
                    }
                    else
                    {
                        Microsoft.TeamFoundation.Client.WindowsCredential wcred = new Microsoft.TeamFoundation.Client.WindowsCredential(true);
                        tfsClientCredentials = new TfsClientCredentials(wcred);
                    }
                }

                using (TfsConfigurationServer serverConfig = new TfsConfigurationServer(uri, tfsClientCredentials))
                {
                    serverConfig.Authenticate();
                    serverConfig.EnsureAuthenticated();
                    if (serverConfig.HasAuthenticated)
                    {
                        Credential = serverConfig.Credentials;
                        catalogNodes = serverConfig.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.ProjectCollection }, false, CatalogQueryOptions.None).OrderBy(f => f.Resource.DisplayName).ToList();
                    }
                }
            }
            catch (TeamFoundationServiceUnavailableException ex)
            {
                MessageBox.Show(ResourceHelper.GetResourceString("MSG_TFS_SERVER_IS_INACCESSIBLE") + "\n" + uri.OriginalString, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                return catalogNodes;
            }

            return catalogNodes;
        }

        #endregion
    }
}
