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
    /// Interaction logic for Options.xaml
    /// </summary>
    public partial class Settings : Window
    {
        public bool m_LinkDocs { get; set; }

        public Settings()
        {            
            InitializeComponent();            
        }

        public void LoadSettingsToForm()
        {
            chkLinkSharedDocs.IsChecked = m_LinkDocs;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            m_LinkDocs = (bool)chkLinkSharedDocs.IsChecked;

            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
