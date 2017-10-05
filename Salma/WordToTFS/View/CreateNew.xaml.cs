using System;
using System.ComponentModel;
using System.Windows;
using WordToTFS;

namespace WordToTFSWordAddIn.Views
{
    public partial class CreateNew : Window
    {
        public System.Windows.Controls.ListBox ListBox { set { listBox1 = value; } get { return listBox1; } }

        //public Action Close;
        public CreateNew()
        {
            InitializeComponent();
        }

        public bool isCancelled;
        public bool isCreated;

        private void CreateButtonClick(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrWhiteSpace(titleTextBox.Text))
            {
                MessageBox.Show(ResourceHelper.GetResourceString("MSG_TITLE_IS_EMPTY"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                isCreated = true;
                Close();
            }
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
