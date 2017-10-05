using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.Integration;
using WordToTFS;

namespace WordToTFSWordAddIn.Views
{
    public partial class ImportDialog : Window
    {
        private PickWorkItemsControl pc = null;

        public bool isCancelled;
        public bool isOk;
        public bool isImportEmptyContent;

        public ImportDialog()
        {
            InitializeComponent();
        }

        public void InsertWorkItemPicker(WorkItemStore WiStore, string ProjectName, bool multiselect)
        {
            pc = new PickWorkItemsControl(WiStore, multiselect);

            if (!ProjectName.Equals(string.Empty))
                pc.PortfolioDisplayName = ProjectName;

            pc.Dock = System.Windows.Forms.DockStyle.Fill;
            pc.AutoSize = true;
            windowsFormsHost1.Child = pc;
        }

        public List<WorkItem> GetSelectedWI()
        {
            if (pc != null)
                return pc.SelectedWorkItems();
            else
                return new List<WorkItem>();
        }

        public void DataWindowClosed(object sender, CancelEventArgs e)
        {
            isCancelled = true;
        }

        private void importButton_Click(object sender, RoutedEventArgs e)
        {
            if (pc == null || pc.SelectedWorkItems().Count == 0)
                MessageBox.Show(ResourceHelper.GetResourceString("SELECTWORKITEMS"));
            else
            {
                isOk = true;

                if (this.contentImportCheckBox.IsChecked == true)
                    isImportEmptyContent = true;

                Close();
            }
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            isCancelled = true;
            Close();
        }
    }
}
