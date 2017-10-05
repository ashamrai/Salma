using System.Windows;
using System.Windows.Controls;

namespace WordToTFS.View
{
    /// <summary>
    /// Interaction logic for StepActionsResulte.xaml
    /// </summary>
    public partial class StepActionsResult : Window
    {
        public bool IsCanceled { set; get; }
        public StepActionsResult()
        {
            InitializeComponent();
            IsCanceled = true;
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            Close();
        }

        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            IsCanceled = false;
            Close();
        }

        public TextBox Action { set { action = value; } get { return action; } }
        public TextBox ExpectedResult { set { expectedResult = value; } get { return expectedResult; } }
    }
}
