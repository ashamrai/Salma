using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using WordToTFS;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace WordToTFSWordAddIn.Views
{
    /// <summary>
    /// Interaction logic for LinkItem.xaml
    /// </summary>
    public partial class LinkItem : Window
    {
        public bool IsCanceled { set; get; }
        public bool IsLink { set; get; }
        public TfsManager Manager { get; set; }
        public List<WorkItem> WorkItemsToLink = new List<WorkItem>();
        private readonly Timer keyEntryTimer;
        

        public LinkItem()
        {
            InitializeComponent();
            IsCanceled = true;
            keyEntryTimer = new Timer(UpdateItems, null, -1, -1);
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            Close();
        }

        private void BtnLinkItem_Click(object sender, RoutedEventArgs e)
        {
            IsCanceled = false;

            if (existingDocumentTabItem.IsSelected)
            {
                if (workItemsListBox.SelectedItems.Count == 0)
                {
                    MessageBox.Show(ResourceHelper.GetResourceString("MSG_NOT_SELECTED_WORK_ITEM"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    Close();
                }
            }
            if (outsideDocumentTabItem.IsSelected)
            {
                if (WorkItemsToLink.Count == 0)
                {
                    MessageBox.Show(ResourceHelper.GetResourceString("MSG_NOT_SELECTED_WORK_ITEM"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    Close();
                }
            }
            if (hyperLinkTabItem.IsSelected)
            {
                if (!string.IsNullOrWhiteSpace(hyperlinkTextBox.Text) && hyperlinkTextBox.Text != "http://")
                {
                    Close();
                }
                else 
                {
                    MessageBox.Show(ResourceHelper.GetResourceString("MSG_CORRECT_HYPERLINK"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
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
            var externalWiIdsText = string.Empty;
            WorkItemsToLink.Clear();
            Dispatcher.Invoke(
                (Action)delegate()
                    {

                        foundItemsListView.Items.Clear();
                        externalWiIdsText = wiIDsTextBox.Text;

                    });


            var itemsToLink = externalWiIdsText.Split(',');
            foreach (var strId in itemsToLink)
            {
                if (!String.IsNullOrWhiteSpace(strId))
                {
                    int id;
                    if (int.TryParse(strId, out id))
                    {
                        var wItem = Manager.GetWorkItem(id);

                        if (wItem != null)
                        {
                            Dispatcher.BeginInvoke((Action)delegate() { foundItemsListView.Items.Add(String.Format("• {0} {1} ({2}): {3} ", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title)); });
                            WorkItemsToLink.Add(wItem);
                        }
                        else
                        {
                            Dispatcher.Invoke((Action)delegate()
                            {
                                foundItemsListView.Items.Add(new ListViewItem()
                                {
                                    Content = String.Format(ResourceHelper.GetResourceString("MSG_ITEM_IS_NOT_FOUND"), id),
                                    Background = new SolidColorBrush(Colors.LightCoral)
                                });

                            });
                        }
                    }
                    else
                    {   
                        Dispatcher.Invoke((Action)delegate()
                        {
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = String.Format(ResourceHelper.GetResourceString("MSG_INPUT_VALUE_INCORRECT"), strId),
                                Background = new SolidColorBrush(Colors.LightCoral)
                            });

                        });
                    }
                }

            }
        }

    }
}



