using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using WordToTFS;

namespace WordToTFSWordAddIn.Views
{
    public partial class UpdateDialog : Window
    {
        public UpdateDialog()
        {
            InitializeComponent();
        }

        public bool isCancelled;
        public bool isOk;
        public bool isUpdateContent;

        private void OkButtonClick(object sender, RoutedEventArgs e)
        {
            isOk = true;

            if (this.contentImportCheckBox.IsChecked == true)
                 isUpdateContent = true;

            Close();
        }

        public void DataWindowClosed(object sender, CancelEventArgs e)
        {
            isCancelled = true;
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            isCancelled = true;
            Close();
        }
    }
}
