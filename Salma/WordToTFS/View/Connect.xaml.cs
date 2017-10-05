using System;
using System.Windows;
using System.Windows.Controls;
using WordToTFS;

namespace WordToTFSWordAddIn.Views
{
    public partial class Connect : Window
    {
        public bool IsCanceled { set; get; }
        public Connect()
        {
            InitializeComponent();
        }


        private void Button1Click(object sender, RoutedEventArgs e)
        {
            IsCanceled = false;
            Close();
        }

        private void Button2Click(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            Close();
        }

        public string Login { set { loginNameBox.Text = value; } get { return loginNameBox.Text; } }
        public string Password { set { passwordBox.Password = value; } get { return passwordBox.Password; } }
    }
}
