using System;
using System.Windows;
using System.Windows.Controls;
using WordToTFS;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Threading;
using System.Windows.Media;
using MessageBox = System.Windows.MessageBox;
using System.Text.RegularExpressions;

namespace WordToTFSWordAddIn.Views
{
    public partial class AddDetails : Window
    {
        public ListBox ListBox { set { listBox1 = value; } get { return listBox1; } }
        public ComboBox FilterBox { set { filterBox = value; } get { return filterBox; } }
        public ComboBox AddDetailsAsBox { set { comboAddDetailsAs = value; } get { return comboAddDetailsAs; } }
        public TextBox GetWIID { set { workItemID = value; } get { return workItemID; } }
        public ListBox FoundItemsListView { set { foundItemsListView = value; } get { return foundItemsListView; } }
        public TextBox WorkItemID { set { workItemID = value; } get { return workItemID; } }

        private string externalWiIdsText = string.Empty;
        bool correctValueInserted = false;
        public List<WorkItem> WorkItemsToLink = new List<WorkItem>();
        private readonly Timer keyEntryTimer;

        public AddDetails()
        {
            InitializeComponent();
            keyEntryTimer = new Timer(UpdateItems, null, -1, -1);
        }

        public bool IsCanceled { set; get; }
        public bool IsAdd { set; get; }
        public bool IsReplace { set; get; }
        public bool IsEmpty { get; set; }

        private delegate void Counter(int i);

        private void AddButtonClick(object sender, RoutedEventArgs e)
        {
            this.IsCanceled = false;
            IsAdd = true;
            IsReplace = false;
            Close();
        }

        private void ReplaceButtonClick(object sender, RoutedEventArgs e)
        {
            IsReplace = true;
            IsCanceled = false;
            IsAdd = false;
            Close();
        }

        public bool ByWorkItemIDTabItem()
        {
            return byWorkItemIDTabItem.IsSelected;
        }

        public bool CurrentDocumentTabItem()
        {
            return currentDocumentTabItem.IsSelected;
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            IsReplace = false;
            IsAdd = false;
            Close();
        }

        public void UpdateItems(object state)
        {
            Dispatcher.Invoke((Action)delegate()
            {
                WorkItemsToLink.Clear();
                foundItemsListView.Items.Clear();
            });

            int itemId = 0;

            if (string.IsNullOrEmpty(externalWiIdsText))
            {
                IsEmpty = true;
            }
            else
            {
                IsEmpty = false;
            }

            if (!String.IsNullOrWhiteSpace(externalWiIdsText))
            {
                int id;
                if (int.TryParse(externalWiIdsText, out id))
                {
                    itemId = id;
                }
                else
                {
                    Dispatcher.BeginInvoke((Action)delegate()
                    {
                        foundItemsListView.Items.Add(new ListViewItem()
                        {
                            Content = String.Format(ResourceHelper.GetResourceString("MSG_INPUT_VALUE_INCORRECT"), externalWiIdsText),
                            Background = new SolidColorBrush(Colors.LightCoral)
                        });
                        addButton.IsEnabled = false;
                        replaceButton.IsEnabled = false;
                        AddDetailsAsBox.IsEnabled = false;
                        AddDetailsAsBox.ItemsSource = null;
                        correctValueInserted = false;
                    });
                }
            }

            WorkItem wItem = TfsManager.Instance.GetWorkItem(itemId);
            if (wItem != null)
            {
                Dispatcher.BeginInvoke((Action)delegate()
                {
                    foundItemsListView.Items.Add(String.Format("• {0} {1} ({2}): {3} ", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title));
                    WorkItemsToLink.Add(wItem);
                });

            }
            else
            {
                if (!IsEmpty)
                {
                    Dispatcher.BeginInvoke((Action)delegate()
                    {
                        int id;
                        if (int.TryParse(externalWiIdsText, out id))
                        {
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = String.Format(ResourceHelper.GetResourceString("MSG_ITEM_IS_NOT_FOUND"), itemId),
                                Background = new SolidColorBrush(Colors.LightCoral)
                            });
                        }
                    });
                }
            }
        }

        /// <summary>
        /// Text Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void workItemID_TextChanged(object sender, TextChangedEventArgs e)
        {
            externalWiIdsText = workItemID.Text;

            if (string.IsNullOrEmpty(externalWiIdsText))
                IsEmpty = true;

            AddDetailsAsBox.IsEnabled = true;
            int wi;

            if (int.TryParse(workItemID.Text, out wi))
            {
                AddDetailsAsBox.ItemsSource = TfsManager.Instance.GetHtmlFieldsByItemId(wi);
                string defaultfield = TfsManager.Instance.GetDefaultDetailsFieldName(wi);

                if (defaultfield != null)
                {
                    AddDetailsAsBox.SelectedValue = defaultfield;

                    string description = TfsManager.Instance.GetWorkItemDescription(wi, defaultfield);

                    if (description == string.Empty)
                    {
                        addButton.IsEnabled = true;
                        replaceButton.IsEnabled = false;
                    }
                    else
                    {
                        addButton.IsEnabled = false;
                        replaceButton.IsEnabled = true;
                    }

                    if (AddDetailsAsBox.SelectedValue.ToString().Equals("Steps") || AddDetailsAsBox.SelectedValue.ToString().Equals("Шаги"))
                    {
                        addButton.IsEnabled = true;
                        replaceButton.IsEnabled = false;
                    }

                    correctValueInserted = true;
                }
                else
                {
                    addButton.IsEnabled = false;
                    replaceButton.IsEnabled = false;
                    AddDetailsAsBox.IsEnabled = false;
                    AddDetailsAsBox.ItemsSource = null;
                    correctValueInserted = false;
                }
                /*
                AddDetailsAsBox.SelectedValue = defaultfield;
                if (AddDetailsAsBox.SelectedValue == "Steps" || AddDetailsAsBox.SelectedValue == "Шаги")
                {
                    replaceButton.IsEnabled = false;
                }
                addButton.IsEnabled = true;
                replaceButton.IsEnabled = true;
                correctValueInserted = true;*/
            }

            keyEntryTimer.Change(500, -1);
        }

        private void listBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            addButton.IsEnabled = true;
            replaceButton.IsEnabled = true;
        }

        // property

        private async void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tabControl = sender as TabControl;

            if (tabControl != null)
            {
                switch (tabControl.SelectedIndex)
                {
                    case 0:
                        if (ListBox.SelectedValue != null)
                        {
                            Match match = Regex.Match(ListBox.SelectedValue.ToString(), @"\d+");
                            int id = 0;

                            if (match.Success && Int32.TryParse(match.Value, out id))
                            {
                                List<object> res = await GetFieldsAsync(id);
                                GetFieldsAsyncCompleted(res);
                                return;
                            }
                        }

                        break;
                    case 1:
                        if (correctValueInserted)
                        {
                            int id;
                            if (int.TryParse(workItemID.Text, out id))
                            {
                                List<object> res = await GetFieldsAsync(id);
                                GetFieldsAsyncCompleted(res);
                                return;
                            }
                        }
                        break;
                }
            }

            addButton.IsEnabled = false;
            replaceButton.IsEnabled = false;
            AddDetailsAsBox.ItemsSource = null;
        }

        private async System.Threading.Tasks.Task<List<object>> GetFieldsAsync(int id)
        {
            List<object> res = new List<object>();
            res.Add(id);

            await System.Threading.Tasks.Task.Run(() =>
            {
                res.Add(TfsManager.Instance.GetHtmlFieldsByItemId(id));
            });
            return res;
        }

        private void GetFieldsAsyncCompleted(List<object> res)
        {
            try
            {
                int id = (int)res[0];
                List<string> items = (List<string>)res[1];

                AddDetailsAsBox.ItemsSource = (List<string>)res[1];
                string defaultField = TfsManager.Instance.GetDefaultDetailsFieldName(id);
                AddDetailsAsBox.SelectedValue = defaultField;

                if (AddDetailsAsBox.SelectedValue != null)
                {
                    string description = TfsManager.Instance.GetWorkItemDescription(id, defaultField);

                    if (description == string.Empty)
                    {
                        addButton.IsEnabled = true;
                        replaceButton.IsEnabled = false;
                    }
                    else
                    {
                        addButton.IsEnabled = false;
                        replaceButton.IsEnabled = true;
                    }

                    if (AddDetailsAsBox.SelectedValue.ToString().Equals("Steps") || AddDetailsAsBox.SelectedValue.ToString().Equals("Шаги"))
                    {
                        addButton.IsEnabled = true;
                        replaceButton.IsEnabled = false;
                    }
                }
            }
            catch { }
        }

        private void ComboAddDetailsAs_DropDownClosed(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    if (ListBox.SelectedValue != null)
                    {
                        Match match = Regex.Match(ListBox.SelectedValue.ToString(), @"\d+");
                        int id = 0;

                        if (match.Success && Int32.TryParse(match.Value, out id) && AddDetailsAsBox.SelectedValue != null)
                        {
                            string description = TfsManager.Instance.GetWorkItemDescription(id, AddDetailsAsBox.SelectedValue.ToString());

                            if (description == string.Empty)
                            {
                                addButton.IsEnabled = true;
                                replaceButton.IsEnabled = false;
                            }
                            else
                            {
                                addButton.IsEnabled = false;
                                replaceButton.IsEnabled = true;
                            }

                            if (AddDetailsAsBox.SelectedValue.ToString().Equals("Steps") || AddDetailsAsBox.SelectedValue.ToString().Equals("Шаги"))
                            {
                                addButton.IsEnabled = true;
                                replaceButton.IsEnabled = false;
                            }

                            return;
                        }
                    }

                    break;
                case 1:
                    if (correctValueInserted)
                    {
                        int id;
                        if (int.TryParse(workItemID.Text, out id) && AddDetailsAsBox.SelectedValue != null)
                        {
                            string description = TfsManager.Instance.GetWorkItemDescription(id, AddDetailsAsBox.SelectedValue.ToString());

                            if (description == string.Empty)
                            {
                                addButton.IsEnabled = true;
                                replaceButton.IsEnabled = false;
                            }
                            else
                            {
                                addButton.IsEnabled = false;
                                replaceButton.IsEnabled = true;
                            }

                            if (AddDetailsAsBox.SelectedValue.ToString().Equals("Steps") || AddDetailsAsBox.SelectedValue.ToString().Equals("Шаги"))
                            {
                                addButton.IsEnabled = true;
                                replaceButton.IsEnabled = false;
                            }

                            return;
                        }
                    }
                    break;
            }

            AddDetailsAsBox.ItemsSource = null;
            addButton.IsEnabled = false;
            replaceButton.IsEnabled = false;
        }
    }
}
