using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
using WordToTFS;
using System.Threading;
using System.Windows.Input;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Windows.Media;
using System.Reflection;
using System.Windows.Controls.Primitives;

namespace WordToTFSWordAddIn.Views
{
    /// <summary>
    /// Interaction logic for QueryReport.xaml
    /// </summary>
    public partial class QueryReport : Window
    {

        public List<WorkItem> WorkItemsToLink = new List<WorkItem>();
        public TfsManager Manager { get; set; }
        public QueryDefinition QueryDef { get; set; }
        public string Project { get; set; }
        private readonly Timer keyEntryTimer;
        public QueryDefinition queryDef;
        private WorkItem wItem;
        bool tabChenged;
        public static bool correctQuery;
        public DateTime selectedDate=DateTime.Now;
        public QueryReport()
        {
            InitializeComponent();
            IsCanceled = true;
            keyEntryTimer = new Timer(UpdateItems, null, -1, -1);
            Manager = TfsManager.Instance;
            onDate.DisplayDateEnd = DateTime.Now;
        }

        public bool IsCanceled { set; get; }

        /// <summary>
        /// Populate Query Fields
        /// </summary>
        /// <param name="allFields">All Field List for project</param>
        /// <param name="displayNames">Display Field List</param>
        public void PopulateQueryFields(List<FieldDefinition> allFields, List<string> displayNames)
        {
            foreach (var field in allFields)
            {
                if (field.Name == "Title" || field.Name == "Название")
                {
                    tbxHeader.Text = field.Name;
                    tbxHeader.IsEnabled = false;
                }
                else
                {

                    if (field.FieldType == FieldType.Html && field.Name != SalmaConstants.TFS.LOCAL_DATA_SOURCE ||
                    (field.FieldType == FieldType.PlainText && field.Name != SalmaConstants.TFS.PROJECT_SERVER_SYNC_ASSIGMENT_DATA) &&
                    field.Name != SalmaConstants.TFS.LOCAL_DATA_SOURCE)
                    {
                        var cbx = CreateListCheckBoxItem(PropertiesListBody, field.Name, displayNames.Contains(field.Name));
                        PropertiesListBody.Items.Add(cbx);
                    }
                    else
                    {
                        var cbx = CreateListCheckBoxItem(PropertiesList, field.Name, displayNames.Contains(field.Name));
                        PropertiesList.Items.Add(cbx);
                    }

                }
            }
            var cbxAll = CreateListCheckBoxItem(PropertiesListBody, ResourceHelper.GetResourceString("ALL"), false);
            PropertiesListBody.Items.Insert(0, cbxAll);
            cbxAll = CreateListCheckBoxItem(PropertiesList, ResourceHelper.GetResourceString("ALL"), false);
            PropertiesList.Items.Insert(0, cbxAll);
        }

        private CheckBox CreateListCheckBoxItem(ListBox listBox, string name, bool isChecked)
        {
            var cbx = new CheckBox();
            cbx.Content = name;
            cbx.ToolTip = name;
            cbx.MinWidth = 190;
            cbx.IsChecked = isChecked;


            cbx.Click += delegate(object sender, RoutedEventArgs e)
            {
                var selectedCbx = sender as CheckBox;
                if (selectedCbx.Content.ToString() == ResourceHelper.GetResourceString("ALL"))
                {
                    SelectAll(listBox, selectedCbx.IsChecked.Value);
                }
            };

            cbx.Checked += delegate(object sender, RoutedEventArgs e)
            {
                var selectedCbx = sender as CheckBox;
                if (selectedCbx.Content.ToString() != ResourceHelper.GetResourceString("ALL"))
                {
                    var chkSelectAll = (CheckBox)listBox.Items[0];
                    if (listBox.Items.Cast<CheckBox>().Count(c => c.IsChecked.Value == true) >= listBox.Items.Count - 1)
                    {
                        chkSelectAll.IsChecked = true;
                    }
                }
            };

            cbx.Unchecked += delegate(object sender, RoutedEventArgs e)
            {
                var selectedCbx = sender as CheckBox;
                if (selectedCbx.Content.ToString() != ResourceHelper.GetResourceString("ALL"))
                {
                    var chkSelectAll = (CheckBox)listBox.Items[0];
                    chkSelectAll.IsChecked = false;
                }

            };

            return cbx;
        }

        /// <summary>
        /// Select All
        /// </summary>
        /// <param name="cbxItem">cbxItem</param>
        private void SelectAll(ListBox listBox, bool isChecked)
        {
            foreach (CheckBox lbxItem in listBox.Items)
            {
                if (lbxItem.Content.ToString() != "Title")
                {
                    lbxItem.IsChecked = isChecked;
                }
            }
        }

        /// <summary>
        /// Insert Button Click event
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">RoutedEventArgs</param>
        private void InsertButtonClick(object sender, RoutedEventArgs e)
        {

            InsertButton.IsEnabled = true;
            IsCanceled = false;
            Close();
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            Close();
        }

        private void SelectionChanged(object sender, RoutedPropertyChangedEventArgs<Object> e)
        {
            if (!correctQuery)
            {
                PropertiesList.Items.Clear();
                PropertiesListBody.Items.Clear();
                tbxHeader.Text = string.Empty;
            }
        }


        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void wiIDsTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            keyEntryTimer.Change(500, -1);
        }

        private void wiIDsTextBox_TextInput(object sender, TextCompositionEventArgs e)
        {

        }

        public void UpdateItems(object state)
        {
            string externalWiIdsText = string.Empty;
            WorkItemsToLink.Clear();
            Dispatcher.Invoke((Action)delegate()
            {
                externalWiIdsText = wiIDsTextBox.Text;
                foundItemsListView.Items.Clear();
            });
            var itemsToLink = externalWiIdsText.Split(new[] {',',';',' ',});
            List<int> updatedIdsList = new List<int>();
            foreach (var strId in itemsToLink)
            {

                if (!String.IsNullOrWhiteSpace(strId))
                {
                    int id;
                    if (int.TryParse(strId, out id))
                    {
                        updatedIdsList.Add(id);
                    }
                    else
                    {
                        Dispatcher.Invoke((Action)delegate()
                        {
                            InsertButton.IsEnabled = false;
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = String.Format(ResourceHelper.GetResourceString("MSG_INPUT_VALUE_INCORRECT"), strId),
                                Background = new SolidColorBrush(Colors.LightCoral)
                            });

                        });
                    }
                }
            }


            string wiNumbers = string.Empty;
            int wInotInCurrentProject = 0;
            foreach (var itemId in updatedIdsList)
            {

                wItem = Manager.GetWorkItem(itemId);
                string m = string.Empty;
                wiNumbers += itemId + ", ";
                if (wItem != null)
                {
                    Dispatcher.Invoke((Action)delegate()
                    {
                        if (wItem != null && wItem.Project.Name != Project)
                        {
                            var textBlock = new TextBlock();
                            textBlock.Text = String.Format("• {0} {1} ({2}): {3} {4} ", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title, ResourceHelper.GetResourceString("NOT_IN_CURRENT_DOCUMENT"));
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = textBlock,
                                Background = new SolidColorBrush(Colors.LightCoral)
                            });
                            workItemNotCurrentProject.Visibility = System.Windows.Visibility.Visible;
                            wInotInCurrentProject++;
                        }
                        else
                        {
                            //var rowN = new ListViewItem() { Content = String.Format("• {0} {1} ({2}): {3}", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title) };
                            //foundItemsListView.Items.Add(rowN);
                            var textBlock = new TextBlock();
                            textBlock.Text = String.Format("• {0} {1} ({2}): {3}", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title);
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = textBlock
                            });
                            //foundItemsListView.Items.Add(String.Format("• {0} {1} ({2}): {3}", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title));
                        }
                        WorkItemsToLink.Add(wItem);
                    });

                }
                else
                {

                    Dispatcher.Invoke((Action)delegate()
                    {
                        InsertButton.IsEnabled = false;
                        foundItemsListView.Items.Add(new ListViewItem()
                        {
                            Content = String.Format(ResourceHelper.GetResourceString("MSG_ITEM_IS_NOT_FOUND"), itemId),
                            Background = new SolidColorBrush(Colors.LightCoral)
                        });
                    });

                }
            }
            if (wInotInCurrentProject == 0)
                Dispatcher.Invoke((Action)delegate()
                    {
                        workItemNotCurrentProject.Visibility = System.Windows.Visibility.Hidden;
                    });

            if (updatedIdsList.Count == 0)
            {
                Dispatcher.Invoke((Action)delegate()
                {
                    InsertButton.IsEnabled = false;
                    tbxHeader.Text = string.Empty;
                    PropertiesList.Items.Clear();
                    PropertiesListBody.Items.Clear();
                });
            }
            var n = foundItemsListView.Items;
            if (updatedIdsList.Count != 0)
            {

                wiNumbers = wiNumbers.Substring(0, wiNumbers.Length - 2);

                string queryText = string.Format("select [System.Id], [System.WorkItemType], [System.Title], [System.State], [System.AreaPath], [System.IterationPath] {0} from WorkItems where [System.TeamProject] = '{1}' and [System.Id] in ({2}) order by [System.ChangedDate] desc", Manager.tfsVersion == TfsVersion.Tfs2011 ? ", [System.Tags]" : string.Empty, Project, wiNumbers);
                queryDef = new QueryDefinition("Work Items list", queryText);
                if (wItem != null)
                {
                    PopulateQueryFieldsFromQuerydef(queryDef);
                }
            }
        }

        public void PopulateQueryFieldsFromQuerydef(QueryDefinition queryDef)
        {
            Query query = new Query(Manager.ItemsStore, queryDef.QueryText);
            var allFields = Manager.GetAllAvailableWorkItemFieldsForProject(Project);
            List<string> displayFieldList = query.DisplayFieldList.Cast<FieldDefinition>().Select(f => f.Name).ToList();
            if (displayFieldList != null)
            {
                Dispatcher.BeginInvoke((Action)delegate()
                {
                    tbxHeader.Clear();
                    PropertiesList.Items.Clear();
                    PropertiesListBody.Items.Clear();
                    PopulateQueryFields(allFields, displayFieldList);
                    InsertButton.IsEnabled = true;
                });
            }
        }


        private void tabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch ((sender as TabControl).SelectedIndex)
            {
                case 0:
                    tabChenged = true;
                    InsertButton.IsEnabled = false;
                    tbxHeader.Text = string.Empty;
                    PropertiesList.Items.Clear();
                    PropertiesListBody.Items.Clear();
                    if (QueryDef != null && correctQuery)
                    {
                        PopulateQueryFieldsFromQuerydef(QueryDef);
                    }
                    onDate.SelectedDate = selectedDate;
                    break;

                case 1:
                    if (tabChenged)
                    {
                        InsertButton.IsEnabled = false;
                        tbxHeader.Text = string.Empty;
                        PropertiesList.Items.Clear();
                        PropertiesListBody.Items.Clear();
                        foundItemsListView.Items.Clear();
                        UpdateItems(wItem);
                        tabChenged = false;
                        onDate.SelectedDate = selectedDate;
                    }
                    break;
            }
        }

        private void onDate_Loaded(object sender, RoutedEventArgs e)
        {
            var picker = sender as DatePicker;
            FieldInfo fiTextBox = picker.GetType().GetField("_textBox", BindingFlags.Instance | BindingFlags.NonPublic);

            if (fiTextBox != null)
            {
                var dateTextBox = (DatePickerTextBox)fiTextBox.GetValue(picker);

                if (dateTextBox != null)
                {
                    PropertyInfo piWatermark = dateTextBox.GetType().GetProperty("Watermark", BindingFlags.Instance | BindingFlags.NonPublic);

                    if (piWatermark != null)
                    {
                        piWatermark.SetValue(dateTextBox, DateTime.Now.ToShortDateString(), null);
                    }
                }
            }
        }

        private void onDate_Selected(object sender, RoutedEventArgs e)
        {
            if (onDate.SelectedDate > DateTime.Now)
            {
                onDate.SelectedDate = DateTime.Now;
            }
            selectedDate = onDate.SelectedDate ?? DateTime.Now;
        }
    }
}

