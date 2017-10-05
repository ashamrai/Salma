using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace WordToTFS.View
{
    public class Win : Window
    {
        public static readonly DependencyProperty TitleProperty = DependencyProperty.Register("Title", typeof(string), typeof(Win), new PropertyMetadata(default(string)));

        public Win()
        {
            InitResources();
        }

        private void InitResources()
        {
            var culture = "en-US";

            switch (Thread.CurrentThread.CurrentUICulture.Name)
            {
                case "ru-RU":
                case "en-US":
                    culture = Thread.CurrentThread.CurrentUICulture.Name;
                    break;
            }


            var res = Application.LoadComponent(
                    new Uri("/WordToTFS;component/View/Localization/StringResources." + culture + ".xaml",
                    UriKind.RelativeOrAbsolute)) as ResourceDictionary;

            Resources.MergedDictionaries.Add(res);
        }

        

        public string Title
        {
            get { return (string)GetValue(TitleProperty); }
            set { SetValue(TitleProperty, value); }
        }

        public static readonly DependencyProperty IconProperty = DependencyProperty.Register("Icon", typeof(Image), typeof(Win), new PropertyMetadata(default(Image)));

        public Image Icon
        {
            get { return (Image)GetValue(IconProperty); }
            set { SetValue(IconProperty, value); }
        }

        public static readonly DependencyProperty CloseButtonProperty = DependencyProperty.Register("CloseButton", typeof(Button), typeof(Win), new PropertyMetadata(default(Button)));

        public Button CloseButton
        {
            get { 
                return (Button)GetValue(CloseButtonProperty); }
            set { SetValue(CloseButtonProperty, value); }
        }
    }
}
