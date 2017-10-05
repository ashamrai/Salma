using System;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using WordToTFS;
using WordToTFS.Model;
using WordToTFS.ViewModel.CreateNew;
using WordToTFSWordAddIn.Views;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Net.Mime;
using System.Security.Cryptography;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Media3D;
using Microsoft.SqlServer.Server;
using WordToTFS.ConfigHelpers;
using System.Runtime.InteropServices;
using MOTW = Microsoft.Office.Tools.Word;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Salma2010
{
    /// <summary>
    /// Add-In
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Current project
        /// </summary>
        public string Project { get; set; }

        /// <summary>
        /// Current area
        /// </summary>
        public string Area { get; set; }

        /// <summary>
        /// Current Link end
        /// </summary>
        public string LinkEnd { get; set; }

        /// <summary>
        /// Current Link WorkItemId
        /// </summary>
        public int LinkWorkItemId { get; set; }

        /// <summary>
        /// Current team project collection
        /// </summary>
        public string TeamProjectCollectionName { get; set; }
        /// <summary>
        /// Work item Id
        /// </summary>
        private int WorkItemId { get; set; }
        /// <summary>
        /// Work item type
        /// </summary>
        private string WorkItemType { get; set; }
        /// <summary>
        /// Gets the ms word version
        /// </summary>
        public MsWordVersion MsWordVersion { get; private set; }

        private List<Microsoft.Office.Tools.CustomTaskPane> WIListTaskPanes = new List<Microsoft.Office.Tools.CustomTaskPane>();
        Microsoft.Office.Tools.CustomTaskPane ActiveTaskPane;

        /// <summary>
        /// Create ribbon
        /// </summary>
        /// <returns></returns>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var app = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            var ci = new CultureInfo((int)app.Language);
            Thread.CurrentThread.CurrentUICulture = ci;

            MsWordVersion = OfficeHelper.GetMsWordVersion(app.Version);
            SectionManager.SetSection(MsWordVersion);
            return new SalmaRibbon();  
        }

        #region VSTO generated code

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            WindowFactory.IconConverter = OleInterop.GetMsoImage;
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
        }


        void Application_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            foreach(Microsoft.Office.Tools.CustomTaskPane _tmpctrl in  WIListTaskPanes)
            {
                WorkItemListControl _wictrl = (WorkItemListControl)_tmpctrl.Control;
                if (_wictrl.DocumentPath == Doc.Path + "\\" + Doc.Name)
                {
                    if (Doc.Name == this.Application.ActiveDocument.Name) ActiveTaskPane = null;
                    WIListTaskPanes.Remove(_tmpctrl);
                    _tmpctrl.Visible = false;
                    _tmpctrl.Dispose();
                    break;
                }
            }
        }

        void Application_WindowActivate(Document Doc, Window Wn)
        {
            ActiveTaskPane = GetActiveTaskPane();

            if (ActiveTaskPane == null) ActiveTaskPane = AddNewTaskPane();
         }

        private Microsoft.Office.Tools.CustomTaskPane AddNewTaskPane()
        {
            if (this.Application.ActiveDocument.Path == "") return null;

            foreach (Microsoft.Office.Tools.CustomTaskPane _tmpctrl in WIListTaskPanes)
            {
                WorkItemListControl _twictrl = (WorkItemListControl)_tmpctrl.Control;
                if (_twictrl.DocumentPath == this.Application.ActiveDocument.Path + "\\" + this.Application.ActiveDocument.Name)
                    return _tmpctrl;
            }

            WorkItemListControl _wictrl = new WorkItemListControl();
            _wictrl.DocumentPath = this.Application.ActiveDocument.Path + "\\" + this.Application.ActiveDocument.Name;
            _wictrl.wordaddin = this;
            _wictrl.SyncWorkItems();
            Microsoft.Office.Tools.CustomTaskPane _witskpane = this.CustomTaskPanes.Add(_wictrl, "Work Item List");
            WIListTaskPanes.Add(_witskpane);
            _witskpane.Visible = false;            
            return _witskpane;
        }

        private Microsoft.Office.Tools.CustomTaskPane GetActiveTaskPane()
        {
            Microsoft.Office.Tools.CustomTaskPane _tskctrl = null;

            foreach (Microsoft.Office.Tools.CustomTaskPane _tmpctrl in WIListTaskPanes)
            {
                WorkItemListControl _wictrl = (WorkItemListControl)_tmpctrl.Control;
                if (_wictrl.DocumentPath == this.Application.ActiveDocument.Path + "\\" + this.Application.ActiveDocument.Name)
                    _tskctrl = _tmpctrl;
            }
            return _tskctrl;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        #region create and add details

        /// <summary>
        /// Add work item
        /// </summary>
        /// <param name="projectName"></param>
        public void AddWorkItem(string projectName, string areapath = "", string linkend = "", int linkid = 0)
        {
            ListBoxItems listBoxCollection = new ListBoxItems(Application.Selection.Text, TfsManager.Instance.GetWorkItemsTypeForCurrentProject(projectName), WorkItemType);

            CreateNew popup = new CreateNew();
            popup.DataContext = listBoxCollection;
            popup.ListBox.ScrollIntoView(WorkItemType);
            popup.Create(null, Icons.AddNewWorkItem);

            if (popup.isCancelled && !popup.isCreated)
                return;

            string errorMessage = string.Empty;

            string type = listBoxCollection.GetValue();
            string title = listBoxCollection.GetTitle();
            int workItemId = CreateNewWi.AddWorkItemForCurrentProject(projectName, title, type, areapath, linkend, linkid);

            WorkItemType = type;

            if (workItemId != 0)
            {
                 WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                 progressDialog.Execute(cancelTokenSource =>
                 {
                     progressDialog.UpdateProgress(100, string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), 1, 1, string.Format("{0} {1}", type, workItemId)), true);

                     try
                     {
                         AddWIControl(workItemId, "System.Title");
                         WorkItemId = workItemId;
                     }
                     catch
                     {
                         errorMessage = ResourceHelper.GetResourceString("MSG_ERROR_ADD_WI");
                     }
                 });

                 progressDialog.Create(ResourceHelper.GetResourceString("MSG_CREATE_WI_TITLE"), Icons.AddNewWorkItem);
            }
            else
                errorMessage = ResourceHelper.GetResourceString("MSG_ERROR_ADD_WI");

            if (!string.IsNullOrEmpty(errorMessage))
                GenerateErrorMessage(errorMessage, Icons.AddNewWorkItem);
        }

        /// <summary>
        /// Add details
        /// </summary>
        public void AddDetails()
        {
            AddDetails popup = new AddDetails();

            popup.Loaded += popup_Loaded;
            popup.Create(null, Icons.AddDetails);

            if (popup.IsAdd || popup.IsReplace)
            {
                string errorMessage = string.Empty;

                int id = 0;
                string fieldName = popup.AddDetailsAsBox.SelectedValue.ToString();

                // if current document TabItem tab taking data from this tab
                if (popup.CurrentDocumentTabItem())
                {
                    Match match = Regex.Match(popup.ListBox.SelectedValue.ToString(), @"\d+");

                    if (match.Success)
                        Int32.TryParse(match.Value, out id);
                }

                // if By Work Item ID TabItem tab taking data from this tab
                if (popup.ByWorkItemIDTabItem())
                {
                    id = Convert.ToInt32(popup.GetWIID.Text);
                }

                if (id != 0)
                {
                    if (TfsManager.Instance.IsStep(id, fieldName))
                    {
                        TfsManager.Instance.AddStep(id, this.Application.Selection.Text);

                        WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                        progressDialog.Execute(cancelTokenSource =>
                        {
                            progressDialog.UpdateProgress(100, string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), 1, 1, string.Format("{0} - {1}", id, fieldName)), true);

                            try
                            {
                                AddWIControl(id, fieldName);
                            }
                            catch
                            {
                                errorMessage = ResourceHelper.GetResourceString("MSG_ERROR_REPLACE_WI");
                            }
                        });

                        progressDialog.Create(ResourceHelper.GetResourceString("MSG_EDIT_WI_TITLE"), Icons.AddDetails);
                    }
                    else
                    {
                        WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                        progressDialog.Execute(cancelTokenSource =>
                        {
                            progressDialog.UpdateProgress(100, string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), 1, 1, string.Format("{0} - {1}", id, fieldName)), true);

                            try
                            {
                                DocxToHtml.Instance.ToHtml(id, fieldName);
                            }
                            catch
                            {
                                errorMessage = ResourceHelper.GetResourceString("MSG_ERROR_REPLACE_WI");
                            }
                        });

                        progressDialog.Create(ResourceHelper.GetResourceString("MSG_EDIT_WI_TITLE"), Icons.AddDetails);
                    }
                }
                else
                    errorMessage = ResourceHelper.GetResourceString("MSG_ERROR_REPLACE_WI");
              

                if (!string.IsNullOrEmpty(errorMessage))
                    GenerateErrorMessage(errorMessage, Icons.AddDetails);
            }
        }

        /// <summary>
        /// popup_Loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void popup_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            AddDetails popup = sender as AddDetails;

            if (popup == null)
                return;

            List<string> itemsSource = GetAllWorkItems();
            popup.ListBox.ItemsSource = itemsSource;

            if (WorkItemId != 0)
            {
                for (int i = 0; i < itemsSource.Count; i++)
                {
                    if (itemsSource[i].Contains(WorkItemId.ToString()))
                    {
                        popup.ListBox.ScrollIntoView(itemsSource[i]);
                        popup.ListBox.SelectedIndex = i;
                        popup.AddDetailsAsBox.IsEnabled = true;
                    }
                }
            }

            popup.ListBox.Focus();

            List<string> types = TfsManager.Instance.GetWorkItemsTypeForCurrentProject(Project);
            types.Add(ResourceHelper.GetResourceString("ALL"));
            popup.FilterBox.ItemsSource = types.OrderBy(f => f).ToList();
            popup.FilterBox.SelectedValue = ResourceHelper.GetResourceString("ALL");
            popup.FilterBox.SelectionChanged += (s, a) =>
            {
                popup.ListBox.ItemsSource = ((string)a.AddedItems[0] == ResourceHelper.GetResourceString("ALL")) ?
                    itemsSource :
                    itemsSource.Where(t => t.Contains(a.AddedItems[0].ToString())).ToList();
            };

            popup.ListBox.SelectionChanged += (s, a) =>
            {
                if (popup.ListBox.SelectedValue != null)
                {
                    popup.AddDetailsAsBox.IsEnabled = true;
                }
                else
                {
                    popup.AddDetailsAsBox.IsEnabled = false;
                    popup.AddDetailsAsBox.ItemsSource = null;
                }
            };
        }

        #endregion

        #region update and sync

        /// <summary>
        /// Update status and sync
        /// </summary>
        public void UpdateStatusAndSync()
        {
            UpdateDialog popup = new UpdateDialog();
            popup.Create(null, Icons.SyncConnectedTool);

            if (popup.isCancelled && !popup.isOk)
                return;

            if (popup.isUpdateContent)
                UpdateStatusAndContent();
            else
                UpdateStatus();
        }

        /// <summary>
        /// Update status
        /// </summary>
        public void UpdateStatus()
        {
            List<ContentControl> controlsToUpdate = (from control in Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                     select control).ToList<ContentControl>();

            if (controlsToUpdate.Count > 0)
            {
                WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                progressDialog.Execute(cancelTokenSource =>
                {
                    progressDialog.UpdateProgress(0, ResourceHelper.GetResourceString("MSG_RETRIEVING_WORK_ITEMS"), true);

                    string errorMessage = string.Empty;
                    int progress = 0;

                    controlsToUpdate.ForEach(control =>
                    {
                        if (!cancelTokenSource.IsCancellationRequested)
                        {
                            progress++;

                            progressDialog.UpdateProgress(Convert.ToInt32(progress / controlsToUpdate.Count * 100),
                                string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), progress, controlsToUpdate.Count, control.Title), true);

                            int id = 0;

                            try
                            {
                                string wiID = this.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;
                                string wiField = this.Application.ActiveDocument.Variables[control.ID + "_wifield"].Value;

                                if (Int32.TryParse(wiID, out id) && TfsManager.Instance.GetWorkItem(id) != null) // Check If comment has link to WI
                                    control.Title = TfsManager.Instance.GetControlText(id, wiField);
                            }
                            catch
                            {
                                if (id > 0)
                                    errorMessage += string.Format("{0}, ", id);
                            }
                        }
                    });

                    if (!string.IsNullOrEmpty(errorMessage))
                        GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_PROCESS_WIS") + errorMessage.Remove(errorMessage.Length - 2), Icons.SyncConnectedTool);
                });

                progressDialog.Create(ResourceHelper.GetResourceString("MSG_UPDATE_WIS_TITLE"), Icons.SyncConnectedTool);
            }
        }

        /// <summary>
        /// Update status and content
        /// </summary>
        public void UpdateStatusAndContent()
        {
            List<ContentControl> controlsToUpdate = (from control in Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                     select control).ToList<ContentControl>();
            if (controlsToUpdate.Count > 0)
            {
                WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                progressDialog.Execute(cancelTokenSource =>
                {
                    progressDialog.UpdateProgress(0, ResourceHelper.GetResourceString("MSG_RETRIEVING_WORK_ITEMS"), true);

                    string errorMessage = string.Empty;
                    int progress = 0;

                    controlsToUpdate.ForEach(control =>
                    {
                        if (!cancelTokenSource.IsCancellationRequested)
                        {
                            int id = 0;
                            string controlTitle = control.Title;

                            progress++;
                            progressDialog.UpdateProgress(Convert.ToInt32(progress / controlsToUpdate.Count * 100),
                                string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), progress, controlsToUpdate.Count, controlTitle), true);

                            try
                            {
                                string wiID = Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;
                                string wiField = Application.ActiveDocument.Variables[control.ID + "_wifield"].Value;
                                int wiRev = 0;
                                int.TryParse(Application.ActiveDocument.Variables[control.ID + "_wirev"].Value, out wiRev);

                                if (!wiField.Equals("Steps", StringComparison.InvariantCultureIgnoreCase) && !wiField.Equals("Шаги", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (Int32.TryParse(wiID, out id))
                                    {
                                        WorkItem _wi = TfsManager.Instance.GetWorkItem(id);
                                        if ( _wi != null)
                                        {
                                            if (wiRev != _wi.Revision)
                                            {
                                                bool historyDescription = true;  // Check if Text Changed in WI Revision // TfsManager.Instance.GetWorkItemFieldHistory(id, wiField, DateTime.Now);

                                                if (historyDescription)
                                                {
                                                    string description = string.Empty;

                                                    //if (wiField.ToLowerInvariant().Equals("System.Title"))
                                                    //    description = _wi.Title;
                                                    //else
                                                        description = _wi.Fields[wiField].Value.ToString();

                                                    HtmlToDocx.Instance.ToDocx(control, description);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                if (id > 0)
                                    errorMessage += string.Format("{0}, ", id);
                            }
                        }
                    });

                    if (!string.IsNullOrEmpty(errorMessage))
                        GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_PROCESS_WIS") + errorMessage.Remove(errorMessage.Length - 2), Icons.SyncConnectedTool);
                });

                progressDialog.Create(ResourceHelper.GetResourceString("MSG_UPDATE_WIS_TITLE"), Icons.SyncConnectedTool);
            }
        }

        #endregion

        #region link, export and import

        /// <summary>
        /// Link work item 
        /// </summary>
        public void LinkItem()
        {
            ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             where Application.Selection.InRange(control.Range)
                                             select control).FirstOrDefault();

            if (contentControl != null)
            {
                int wiID = 0;

                try
                {
                    if (Int32.TryParse(this.Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value, out wiID))
                    {
                        List<string> itemsSource = new List<string>();

                        List<ContentControl> controls = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                         where control.ID != contentControl.ID
                                                         select control).ToList<ContentControl>();
                        
                        controls.ForEach(control =>
                        {
                            try
                            {
                                string wID = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;
                                string wiField = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wifield"].Value;

                                int id = 0;

                                if (int.TryParse(wID, out id))
                                {
                                    Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wItem = TfsManager.Instance.GetWorkItem(id);

                                    if (wItem != null)
                                        itemsSource.Add(String.Format("• {0} {1} - {2} ({3}) : {4} ", wItem.Type.Name, wItem.Id, 
                                            wItem.Fields[wiField].Name, wItem.State, wItem.Title));
                                }
                            }
                            catch { }
                        });

                        LinkWorkItem LinkWorkItem = new LinkWorkItem(Application.ActiveDocument);
                        LinkWorkItem.LinkItem(itemsSource, wiID, Project);
                    }
                        
                }
                catch
                {
                    GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_PROCESS_WI") + wiID, Icons.LinkItems);
                }
            }
        }

        /// <summary>
        /// Export items
        /// </summary>
        public void ExportItem()
        {
            ExportDialog popup = new ExportDialog();
            popup.Create(null, Icons.ExportItem);

            if (popup.isCancelled && !popup.isOk)
                return;

            ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             where Application.Selection.InRange(control.Range)
                                             select control).FirstOrDefault();

            if (contentControl != null)
            {
                WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                progressDialog.Execute(cancelTokenSource =>
                {
                    progressDialog.UpdateProgress(100, string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), 1, 1, contentControl.Title), true);

                    string errorMessage = string.Empty;

                    if (!cancelTokenSource.IsCancellationRequested)
                    {
                        int id = 0;

                        try
                        {
                            string wiID = Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value;
                            string wiField = Application.ActiveDocument.Variables[contentControl.ID + "_wifield"].Value;

                            if (Int32.TryParse(wiID, out id) && TfsManager.Instance.GetWorkItem(id) != null) // Check If comment has link to WI
                            {
                                contentControl.Range.Select();

                                if (wiField.Equals("System.Title", StringComparison.InvariantCultureIgnoreCase))
                                    TfsManager.Instance.ReplaceDetailsForWorkItem(id, wiField, Application.Selection.Text);
                                else
                                {
                                    DocxToHtml.Instance.ToHtmlExport(id, wiField);
                                }
                            }
                        }
                        catch
                        {
                            if (id > 0)
                                errorMessage += id;
                        }
                    }

                    Application.Selection.Collapse();

                    if (!string.IsNullOrEmpty(errorMessage))
                        GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_EXPORT_WI") + errorMessage, Icons.ExportItem);
                });

                progressDialog.Create(ResourceHelper.GetResourceString("MSG_EXPORT_WI_TITLE"), Icons.ExportItem);
            }
        }

        /// <summary>
        /// Import items
        /// </summary>
        public void ImportItems()
        {

            ImportDialog popup = new ImportDialog();
            popup.InsertWorkItemPicker(TfsManager.Instance.ItemsStore, this.Project, true);
            popup.Create(null, Icons.ImportItems);

            if (popup.isCancelled && !popup.isOk)
                return;

            List<Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem> items = popup.GetSelectedWI();
            bool isImportEmptyContent = popup.isImportEmptyContent;

            if (items.Count > 0)
            {
                Application.Selection.Paragraphs.Last.Range.Paragraphs.Add();
                Range rangeToInsert = Application.Selection.Paragraphs.Last.Range;

                List<int> ids = GetAllWorkItemsIds();

                WordToTFS.View.ProgressDialog progressDialog = new WordToTFS.View.ProgressDialog();

                progressDialog.Execute(cancelTokenSource =>
                {
                    string errorMessage = string.Empty;
                    int progress = 0;

                    progressDialog.UpdateProgress(0, ResourceHelper.GetResourceString("MSG_RETRIEVING_WORK_ITEMS"), true);

                    items.ForEach(item =>
                    {
                        if (!cancelTokenSource.IsCancellationRequested && !ids.Contains(item.Id))
                        {
                            try
                            {
                                progress++;
                                progressDialog.UpdateProgress(Convert.ToInt32(progress / items.Count * 100),
                                    string.Format(ResourceHelper.GetResourceString("MSG_PROGRESS"), progress, items.Count, string.Format("{0} {1}", item.Type.Name, item.Id)), true);

                                Globals.ThisAddIn.Application.ScreenUpdating = false;

                                Paragraph p = rangeToInsert.Paragraphs.Add();

                                p.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                                p.Range.Text = item.Title;
                                p.Range.Select();
                                Application.Selection.ClearFormatting();

                                AddWIControl(item.Id, "System.Title");
                                p.Range.InsertParagraphAfter();
                                rangeToInsert = p.Range;
                                
                                if (!string.IsNullOrEmpty(item.Description))
                                {
                                    p = rangeToInsert.Paragraphs.Add();
                                    HtmlToDocx.Instance.ToDocxImport(p.Range, item.Description, item.Id, item.Fields["System.Description"].Name);
                                    p.Range.InsertParagraphAfter();
                                    rangeToInsert = p.Range;
                                }
                                else if (isImportEmptyContent)
                                {
                                    p = rangeToInsert.Paragraphs.Add();
                                    p.Range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                                    p.Range.Select();
                                    Application.Selection.ClearFormatting();
                                    AddWIControl(item.Id, item.Fields["System.Description"].Name);
                                    rangeToInsert = p.Range;
                                }

                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                            }
                            catch
                            {
                                errorMessage += string.Format("{0}, ", item.Id);
                            }
                        }
                    });

                    Application.Selection.Collapse();

                    if (!string.IsNullOrEmpty(errorMessage))
                        GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_IMPORT_WIS") + errorMessage.Remove(errorMessage.Length - 2), Icons.ExportItem);
                });

                progressDialog.Create(ResourceHelper.GetResourceString("MSG_IMPORT_WIS_TITLE"), Icons.ImportItems);
            }
        }

        #endregion

        #region content control

        /// <summary>
        /// Get work items
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllWorkItems()
        {
            List<string> wItems = new List<string>();

            List<ContentControl> controls = (from control in Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                     select control).ToList<ContentControl>();

            controls.ForEach(control =>
            {
                try
                {
                    string wID = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;
                    string wiField = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wifield"].Value;

                    int id;

                    if (int.TryParse(wID, out id))
                    {
                        Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wItem = TfsManager.Instance.GetWorkItem(id);

                        if (wItem != null)
                            wItems.Add(String.Format("• {0} {1} - {2} ({3}) : {4} ", wItem.Type.Name, wItem.Id, wItem.Fields[wiField].Name, wItem.State, wItem.Title));
                    }
                }
                catch { }
            });

            return wItems;
        }

        /// <summary>
        /// Get work items ids
        /// </summary>
        /// <returns></returns>
        public List<int> GetAllWorkItemsIds()
        {
            List<int> wItems = new List<int>();

            List<ContentControl> controls = (from control in Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             select control).ToList<ContentControl>();

            controls.ForEach(control =>
            {
                try
                {
                    string wID = Globals.ThisAddIn.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value;

                    int id;

                    if (int.TryParse(wID, out id) && TfsManager.Instance.GetWorkItem(id) != null)
                        wItems.Add(id);
                }
                catch { }
            });

            return wItems;
        }
        /// <summary>
        /// Enable open control description button
        /// </summary>
        /// <returns></returns>
        public bool IsWorkItemInContentControl()
        {
            int count = 0;

            try
            {
                count = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                         where Application.Selection.InRange(control.Range)
                         select control).Count();
            }
            catch(Exception)
            {
                count = 0;
            }

            return count > 0 ? true : false;
        }

        /// <summary>
        /// Open work item
        /// </summary>
        public void OpenWorkItem()
        {
            ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             where Application.Selection.InRange(control.Range)
                                             select control).FirstOrDefault();

            if (contentControl != null)
            {
                try
                {
                    string wiID = this.Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value;

                    int id = 0;

                    if (Int32.TryParse(wiID, out id))
                        System.Diagnostics.Process.Start(TfsManager.Instance.GetWorkItemLink(id, this.Project));
                }
                catch
                {
                    GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_OPEN_WI"), Icons.OpenWorkItem);
                }
            }
        }

        /// <summary>
        /// Add/replace work item control
        /// </summary>
        /// <param name="wiID"></param>
        /// <param name="wiFieldName"></param>
        /// <param name="replace"></paramprivate
        public void AddWIControl(int wiID, string wiFieldName, Range range = null, string OldControlID = "")
        {
            Range selection = null;

            if (range == null)
                selection = this.Application.ActiveDocument.Range(Application.Selection.Range.Start, Application.Selection.Range.End);
            else
                selection = range;

            ClearControls(wiID, wiFieldName);

            // add new control
            ContentControl newControl = this.Application.ActiveDocument.ContentControls.Add(WdContentControlType.wdContentControlRichText, selection);            
            newControl.Title = TfsManager.Instance.GetControlText(wiID, wiFieldName);

            if (this.MsWordVersion == MsWordVersion.MsWord2013)
                     newControl.Color = WdColor.wdColorLightBlue;

            // save control properties
            Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wItem = TfsManager.Instance.GetWorkItem(wiID);
            AddVariables(wItem, wiFieldName, newControl.ID);
            if (OldControlID == "") AddPaneValues(wItem, wiFieldName, newControl.ID);
            else UpdatePaneControlID(OldControlID, newControl.ID);
        }

        private void UpdatePaneControlID(string pOldControlID, string pNewControlID)
        {
            if (ActiveTaskPane == null) return;
            WorkItemListControl _wictrl = (WorkItemListControl)ActiveTaskPane.Control;
            _wictrl.Invoke(_wictrl.UpdateControlIDDelegate, new object[] { pOldControlID, pNewControlID });
        }

        private void AddPaneValues(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem pWi, string wiFieldName, string controlID)
        {
            if (ActiveTaskPane == null) return;
            WorkItemListControl _wictrl = (WorkItemListControl)ActiveTaskPane.Control;
            _wictrl.Invoke(_wictrl.AddWiDelegate, new object[] { pWi.Id.ToString(), pWi.Fields[wiFieldName].Name, pWi.Title, controlID });
        }

        /// <summary>
        /// Add control variables
        /// </summary>
        /// <param name="wiID"></param>
        /// <param name="wiFieldName"></param>
        /// <param name="controlID"></param>
        private void AddVariables(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem pWi, string wiFieldName, string controlID)
        {
            try
            {
                this.Application.ActiveDocument.Variables.Add(controlID + "_wiid", pWi.Id);
                this.Application.ActiveDocument.Variables.Add(controlID + "_wifield", wiFieldName);
                this.Application.ActiveDocument.Variables.Add(controlID + "_wifieldEx", pWi.Fields[wiFieldName].Name);
                this.Application.ActiveDocument.Variables.Add(controlID + "_wititle", pWi.Title);
                this.Application.ActiveDocument.Variables.Add(controlID + "_wirev", pWi.Rev);
            }
            catch { }
        }

        /// <summary>
        /// Clear controls
        /// </summary>
        /// <param name="wiID"></param>
        /// <param name="wiFieldName"></param>
        private void ClearControls(int wiID, string wiFieldName)
        {
            List<ContentControl> controls = (from control in Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             select control).ToList<ContentControl>();

            controls.ForEach( control => 
            {
                try
                {
                    if (this.Application.ActiveDocument.Variables[control.ID + "_wiid"].Value.Equals(wiID.ToString(), StringComparison.InvariantCultureIgnoreCase) &&
                        this.Application.ActiveDocument.Variables[control.ID + "_wifield"].Value.Equals(wiFieldName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        DeleteVariables(control.ID);
                        control.Delete(false);
                    }
                }
                catch { }
            });
        }

        /// <summary>
        /// Delete control variables
        /// </summary>
        /// <param name="controlID"></param>
        public void DeleteVariables(string controlID)
        {
            try { this.Application.ActiveDocument.Variables[controlID + "_wiid"].Delete(); }
            catch { }

            try { this.Application.ActiveDocument.Variables[controlID + "_wifield"].Delete(); }
            catch { }

            try { this.Application.ActiveDocument.Variables[controlID + "_wifieldEx"].Delete(); }
            catch { }

            try { this.Application.ActiveDocument.Variables[controlID + "_wititle"].Delete(); }
            catch { }

            try { this.Application.ActiveDocument.Variables[controlID + "_wirev"].Delete(); }
            catch { }
        }

        /// <summary>
        /// Is step
        /// </summary>
        /// <returns></returns>
        public bool IsStep()
        {
            ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             where Application.Selection.InRange(control.Range)
                                             select control).FirstOrDefault();

            if (contentControl != null)
            {
                try
                {
                    string wiField = this.Application.ActiveDocument.Variables[contentControl.ID + "_wifield"].Value;

                    if (wiField.Equals("Steps", StringComparison.InvariantCultureIgnoreCase) || wiField.Equals("Шаги", StringComparison.InvariantCultureIgnoreCase))
                        return true;
                }
                catch
                {
                }
            }

            return false;
        }
        #endregion

        #region report

        /// <summary>
        /// Generate report
        /// </summary>
        public void GenerateReport()
        {
            var report = new Report(Application.ActiveDocument);
            report.GenerateReport(Project, NormalText);
        }

        /// <summary>
        /// Normal text
        /// </summary>
        /// <param name="text"></param>
        private void NormalText(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }

            Object styleTitle = WdBuiltinStyle.wdStyleNormalObject;
            Object styleTitle2 = WdBuiltinStyle.wdStyleHtmlNormal;
            var doc = Application.ActiveDocument;
            var pt = OfficeHelper.CreateParagraphRange(ref doc);
            pt.Text = text;
            pt.set_Style(ref styleTitle2);

            var thread = new Thread(SetClipboard);
            if (thread.TrySetApartmentState(ApartmentState.STA))
            {
                thread.Start(text);
                thread.Join();
            }

            try
            {
                object objDataTypeMetafile = WdPasteDataType.wdPasteHTML;
                pt.PasteSpecial(DataType: objDataTypeMetafile);
                if (Application.ActiveDocument.InlineShapes.Count > 0)
                {
                    var page = Application.ActiveDocument.PageSetup;
                    float calculatedWidth = page.PageWidth - (page.LeftMargin + page.RightMargin);

                    foreach (InlineShape shape in Application.ActiveDocument.InlineShapes)
                    {
                        shape.LockAspectRatio = MsoTriState.msoTrue;

                        if (shape.Width <= calculatedWidth) continue;

                        shape.Width = calculatedWidth;
                    }
                }
            }
            catch (Exception)
            {
            }

            Application.ActiveDocument.Content.InsertParagraphAfter();
        }

        /// <summary>
        /// Set clipboard
        /// </summary>
        /// <param name="text"></param>
        private void SetClipboard(object text)
        {
            try
            {
                CopyToClipboard((string)text);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Copy to clipboard
        /// </summary>
        /// <param name="html"></param>
        public static void CopyToClipboard(string html)
        {
            Encoding encoding = Encoding.UTF8;

            const int numberLengthWithCr = 11;

            var htmlIntro = "<html>\n<head>\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=" + encoding.WebName + "\" />\n</head>\n<body>\n<!--StartFragment-->";
            string htmlOutro = "<!--EndFragment-->\n</body>\n</html>";

            int startHtmlIndex = 57 + 4 * numberLengthWithCr;
            int startFragmentIndex = startHtmlIndex + encoding.GetByteCount(htmlIntro);
            int endFragmentIndex = startFragmentIndex + encoding.GetByteCount(html);
            int endHtmlIndex = endFragmentIndex + encoding.GetByteCount(htmlOutro);

            StringBuilder buff = new StringBuilder();
            buff.AppendFormat("Version:1.0\n");
            buff.AppendFormat("StartHTML:{0:0000000000}\n", startHtmlIndex);
            buff.AppendFormat("EndHTML:{0:0000000000}\n", endHtmlIndex);
            buff.AppendFormat("StartFragment:{0:0000000000}\n", startFragmentIndex);
            buff.AppendFormat("EndFragment:{0:0000000000}\n", endFragmentIndex);
            buff.Append(htmlIntro).Append(html).Append(htmlOutro);

            ClipboardHelper.Instanse.CopyToClipboard(buff.ToString());
        }

        #endregion

        /// <summary>
        /// Generate matrix
        /// </summary>
        public void GenerateMatrix()
        {
            var matrix = new TraceabilityMatrix(Application.ActiveDocument);
            matrix.GenerateMatrix(Project);
        }

        /// <summary>
        /// Generate error message
        /// </summary>
        /// <param name="message"></param>
        /// <param name="icon"></param>
        private void GenerateErrorMessage(string message, Icons icon)
        {
            ErrorDialog popup = new ErrorDialog();
            popup.MessageBlock.Text = message;
            popup.Create(null, icon);
        }

        public void ShowPanel()
        {
            if (ActiveTaskPane != null)
            {
                if (ActiveTaskPane.Visible) ActiveTaskPane.Visible = false;
                else ActiveTaskPane.Visible = true;
            }    
            else
            {
                ActiveTaskPane = AddNewTaskPane();
            }
        }

        internal void ObsoleteWorkItem(string linkend = "", int linkid = 0)
        {
            ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                             where Application.Selection.InRange(control.Range)
                                             select control).FirstOrDefault();

            if (contentControl != null)
            {
                try
                {
                    string wiID = this.Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value;

                    int id = 0;

                    if (Int32.TryParse(wiID, out id))
                    {
                        List<WordToTFS.WILink> WiLinks = new List<WILink>();

                        Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wItem = TfsManager.Instance.GetWorkItem(id);

                        GetObsoleteWorkItemLinks(WiLinks, wItem);

                        WordToTFS.View.ObsoleteWorkItem ObsoleteWiWindow = new WordToTFS.View.ObsoleteWorkItem(WiLinks);
                        ObsoleteWiWindow.ObsoleteTagChecked = Properties.Settings.Default.Obsolete_Tag_Checked;
                        ObsoleteWiWindow.ObsoleteTitleChecked = Properties.Settings.Default.Obsolete_Title_Checked;
                        ObsoleteWiWindow.ObsoleteTagText = Properties.Settings.Default.Obsolete_Tag_Text;
                        ObsoleteWiWindow.ObsoleteTitleText = Properties.Settings.Default.Obsolete_Title_Text;

                        ObsoleteWiWindow.ShowDialog();

                        if (ObsoleteWiWindow.OperationAccepted)
                        {
                            if (ObsoleteWiWindow.SettingsChanged)
                            {
                                Properties.Settings.Default.Obsolete_Tag_Checked = ObsoleteWiWindow.ObsoleteTagChecked;
                                Properties.Settings.Default.Obsolete_Title_Checked = ObsoleteWiWindow.ObsoleteTitleChecked;
                                Properties.Settings.Default.Obsolete_Tag_Text = ObsoleteWiWindow.ObsoleteTagText;
                                Properties.Settings.Default.Obsolete_Title_Text = ObsoleteWiWindow.ObsoleteTitleText;
                                Properties.Settings.Default.Save();
                            }

                            DuplicateWorkItem(wItem, ObsoleteWiWindow.GetUpdatedLinks(), linkend, linkid);
                        }
                    }
                }
                catch(Exception ex)
                {
                    GenerateErrorMessage(ResourceHelper.GetResourceString("MSG_ERROR_OBSOLETE_WI") + "\n" + ex.StackTrace, Icons.OpenWorkItem);
                }
            }
            
        }

        private void DuplicateWorkItem(WorkItem pWiItem, List<WILink> pWiList, string linkend = "", int linkid = 0)
        {
            List<WordContentControl> WIControls = new List<WordContentControl>();

            foreach(ContentControl contentControl in this.Application.ActiveDocument.ContentControls)
            {
                try
                {
                    if (this.Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value == pWiItem.Id.ToString())
                    {
                        WordContentControl wordContentCtrl = new WordContentControl();
                        wordContentCtrl.ControlId = contentControl.ID;
                        wordContentCtrl.WorkItemId = this.Application.ActiveDocument.Variables[contentControl.ID + "_wiid"].Value;
                        wordContentCtrl.WorkItemField = this.Application.ActiveDocument.Variables[contentControl.ID + "_wifield"].Value;
                        wordContentCtrl.WorkItemFieldEx = this.Application.ActiveDocument.Variables[contentControl.ID + "_wifieldEx"].Value;
                        wordContentCtrl.WorkItemTitle = this.Application.ActiveDocument.Variables[contentControl.ID + "_wititle"].Value;
                        wordContentCtrl.WorkItemRev = this.Application.ActiveDocument.Variables[contentControl.ID + "_wirev"].Value;

                        WIControls.Add(wordContentCtrl);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                }
            }

            if (WIControls.Count > 0)
            {
                bool SourceWiChaged = false;

                WorkItem newItem = pWiItem.Store.Projects[pWiItem.Project.Name].WorkItemTypes[pWiItem.Type.Name].NewWorkItem();

                newItem.AreaPath = pWiItem.AreaPath;
                newItem.Tags = pWiItem.Tags;

                if (Properties.Settings.Default.Obsolete_Tag_Checked)
                {
                    if (pWiItem.Tags != "") pWiItem.Tags += ";" + Properties.Settings.Default.Obsolete_Tag_Text;
                    else pWiItem.Tags = Properties.Settings.Default.Obsolete_Tag_Text;
                    SourceWiChaged = true;
                }

                foreach (WordContentControl wordControl in WIControls)
                {
                    newItem.Fields[wordControl.WorkItemFieldEx].Value = pWiItem.Fields[wordControl.WorkItemFieldEx].Value;

                    if (newItem.Fields[wordControl.WorkItemFieldEx].ReferenceName == "System.Title"
                        && Properties.Settings.Default.Obsolete_Title_Checked)
                    {
                        pWiItem.Fields[wordControl.WorkItemFieldEx].Value = Properties.Settings.Default.Obsolete_Title_Text + " " +
                            pWiItem.Fields[wordControl.WorkItemFieldEx].Value;

                        SourceWiChaged = true;
                    }
                }

                if (linkend != "" && linkid != 0)
                    if (TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds.Contains(linkend))
                        newItem.Links.Add(new RelatedLink(TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds[linkend], linkid));

                foreach(WILink _wiLink in pWiList)
                {
                    if (_wiLink.SelectedActionType == WordToTFS.Properties.Resources.Obolete_List_Item_Move)
                        for (int i = 0; i < pWiItem.WorkItemLinks.Count; i++ )
                            if (pWiItem.WorkItemLinks[i].LinkTypeEnd.Name == _wiLink.LinkTypeEnd
                                && pWiItem.WorkItemLinks[i].TargetId == _wiLink.TargetId)
                            {
                                pWiItem.WorkItemLinks.RemoveAt(i);
                                SourceWiChaged = true;
                                break;
                            }

                    if (TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds.Contains(_wiLink.LinkTypeEnd))
                        newItem.Links.Add(new RelatedLink(TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds[_wiLink.LinkTypeEnd], _wiLink.TargetId));
                }

                if (SourceWiChaged)
                    if (SaveWorkItem(pWiItem) == 0) return;

                var _lknRelated = (from _lnk in TfsManager.Instance.ItemsStore.WorkItemLinkTypes where _lnk.ReferenceName == "System.LinkTypes.Related" select _lnk).FirstOrDefault();

                if (_lknRelated != null)
                {
                    RelatedLink rel = new RelatedLink(_lknRelated.ForwardEnd, pWiItem.Id);
                    rel.Comment = "Copied from " + pWiItem.Id.ToString();
                    newItem.Links.Add(rel);
                }

                int newId = SaveWorkItem(newItem);

                if (newId != 0)
                {
                    foreach (WordContentControl wordControl in WIControls)
                    {
                        DeleteVariables(wordControl.ControlId);
                        AddVariables(newItem, wordControl.WorkItemField, wordControl.ControlId);

                        ContentControl contentControl = (from control in this.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                         where control.ID == wordControl.ControlId
                                                         select control).FirstOrDefault();

                        contentControl.Title = TfsManager.Instance.GetControlText(newId, wordControl.WorkItemField);
                    }

                }
            }
        }

        private static int SaveWorkItem(WorkItem pWiItem)
        {
            int id = 0;

            var Errors = pWiItem.Validate();

            if (Errors.Count > 0)
            {
                
            }
            else
            {
                pWiItem.Save();
                id = pWiItem.Id;
            }

            return id;
        }

        private static void GetObsoleteWorkItemLinks(List<WordToTFS.WILink> WiLinks, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem wItem)
        {
            foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemLink wiLink in wItem.WorkItemLinks)
            {
                Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem linkedItem = TfsManager.Instance.GetWorkItem(wiLink.TargetId);

                WordToTFS.WILink itemLink = new WordToTFS.WILink()
                {
                    TargetId = wiLink.TargetId,
                    LinkTypeEnd = wiLink.LinkTypeEnd.Name,
                    TargetWorkItemType = linkedItem.Type.Name,
                    TargetWorkItemTitle = linkedItem.Title
                };

                if (wiLink.LinkTypeEnd.LinkType.LinkTopology == WorkItemLinkType.Topology.Tree)
                {
                    if (wiLink.LinkTypeEnd.IsForwardLink)
                    {
                        itemLink.IsSelected = false;
                        itemLink.RemoveCopyAction();
                    }
                }

                WiLinks.Add(itemLink);
            }
        }
    }
}