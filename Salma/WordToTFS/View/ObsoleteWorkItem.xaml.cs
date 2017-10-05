using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordToTFS.View
{
    /// <summary>
    /// Логика взаимодействия для ObsoleteWorkItem.xaml
    /// </summary>
    public partial class ObsoleteWorkItem : Window
    {
        public bool OperationAccepted = false;
        public bool SettingsChanged = false;
        public List<WILink> WorkItemLinks = null;
        public bool ObsoleteTagChecked
        {
            get { return (bool)TagSelected.IsChecked; }
            set { TagSelected.IsChecked = value; }
        }
        public bool ObsoleteTitleChecked
        {
            get { return (bool)TitleSelected.IsChecked; }
            set { TitleSelected.IsChecked = value; }
        }
        public string ObsoleteTagText
        {
            get { return txtWorkItemTag.Text; }
            set { txtWorkItemTag.Text = value; }
        }
        public string ObsoleteTitleText
        {
            get { return txtWorkItemTitle.Text; }
            set { txtWorkItemTitle.Text = value; }
        }

        public ObsoleteWorkItem(List<WILink> pWorkItemLinks)
        {
            InitializeComponent();

            WorkItemLinks = pWorkItemLinks;
            Links.ItemsSource = WorkItemLinks;
        }

        public List<WILink> GetUpdatedLinks()
        {
            if (Links.ItemsSource == null) return null;

            return (List<WILink>)Links.ItemsSource;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            OperationAccepted = true;
            this.Close();
        }

        private void OkCancel_Click(object sender, RoutedEventArgs e)
        {
            OperationAccepted = false;
            this.Close();
        }

        private void TitleSelected_Click(object sender, RoutedEventArgs e)
        {
            SettingsChanged = true;
        }

        private void TagSelected_Click(object sender, RoutedEventArgs e)
        {
            SettingsChanged = true;
        }

        private void txtWorkItemTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            SettingsChanged = true;
        }

        private void txtWorkItemTag_TextChanged(object sender, TextChangedEventArgs e)
        {
            SettingsChanged = true;
        }
    }
}
