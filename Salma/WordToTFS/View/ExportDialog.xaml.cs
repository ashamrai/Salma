using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using WordToTFS;

namespace WordToTFSWordAddIn.Views
{
    public partial class ExportDialog : Window
    {
        public ExportDialog()
        {
            InitializeComponent();
        }

        public bool isCancelled;
        public bool isOk;

        private void OkButtonClick(object sender, RoutedEventArgs e)
        {
            isOk = true;
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
