using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace WordToTFS.View
{
    /// <summary>
    /// Interaction logic for ProgressDialog.xaml
    /// </summary>
    public partial class ProgressDialog : Window
    {
        private delegate void UpdateProgressDelegate(int percentage, string text, bool IsIndeterminate);
        private Task task;
        private CancellationTokenSource cancelToken;
        //public Action Close;
        public ProgressDialog()
        {
            InitializeComponent();
            cancelToken = new CancellationTokenSource();

        }

        public void Execute(Action<CancellationTokenSource> action)
        {
            task = new Task(() => action(cancelToken), cancelToken.Token);
            
            task.ContinueWith((t) =>
            {
                this.Dispatcher.BeginInvoke(new Action(()=> Close()));
            });
            task.Start();

            
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            cancelToken.Cancel();
            if (task.IsCanceled)
            {
                Close();
            }
        }

        public void UpdateProgress(int percentage, string msg, bool IsIndeterminate)
        {   
            this.Dispatcher.BeginInvoke(new UpdateProgressDelegate((p, m, isInd) => {
                lblProgress.Text = m;
                progress.Value = p;
                progress.IsIndeterminate = isInd;
            }), percentage, msg, IsIndeterminate);
        }

        protected override void OnClosed(EventArgs e)
        {
            cancelToken.Cancel();
            base.OnClosed(e);
        }
        private void Win_Unloaded(object sender, RoutedEventArgs e)
        {
            cancelToken.Cancel();
        }
    }
}
