using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using WordToTFS.ViewModel;

namespace WordToTFS.View
{
    /// <summary>
    /// Interaction logic for RequiredFields.xaml
    /// </summary>
    public partial class RequiredFields
    {
        public bool IsCancelled = false;
        public bool IsCreated;
        public bool TextBoxChanged = false;
        public RequiredFields()
        {
           
            InitializeComponent();
        }

        private void CreateButton1Click(object sender, RoutedEventArgs e)
        {

            InvalidateArrange();
            string requiredFields = string.Empty;
            string numericFields = string.Empty;
            var requiredFieldsList = (List<IViewModel>) DataContext;
            int RequiredFieldCount = 0;
            int NumericFieldCount = 0;
            int temp = Resources.Count;

            foreach (IViewModel model in requiredFieldsList)
            {
                if (model.GetValue() == null || string.IsNullOrEmpty(model.GetValue().ToString().Trim()))
                {
                    requiredFields += string.Format("{0}, ", model.GetName());
                    RequiredFieldCount++;
                }
                else if (model.IsNumeric)
                {
                    if (!IsInteger((string)model.GetValue()))
                    {
                        numericFields += string.Format("{0}, ", model.GetName());
                        NumericFieldCount++;
                    }
                }

            }
            if (NumericFieldCount != 0)
            {
                if (NumericFieldCount == 1)
                {
                    MessageBox.Show(string.Format("{0} : {1} {2}",
                                                  WordToTFS.Properties.Resources.RequiredFields_Fields_singular,
                                                  numericFields.Remove(numericFields.Length - 2, 2),
                                                  WordToTFS.Properties.Resources
                                                          .RequiredFields_Must_Contain_Numeric_Value_singular));
                    return;
                }
                MessageBox.Show(string.Format("{0} : {1} {2}",
                                              WordToTFS.Properties.Resources.RequiredFields_Fields_plural,
                                              numericFields.Remove(numericFields.Length - 2, 2),
                                              WordToTFS.Properties.Resources
                                                      .RequiredFields_Must_Contain_Numeric_Value_plural));
                return;
            }
            if (RequiredFieldCount != 0)
            {
                if (RequiredFieldCount == 1)
                {
                    MessageBox.Show(string.Format("{0} : {1} {2}",
                                                  WordToTFS.Properties.Resources.RequiredFields_Fields_singular,
                                                  requiredFields.Remove(requiredFields.Length - 2, 2),
                                                  WordToTFS.Properties.Resources.RequiredFields_Required_singular));
                    return;
                }
                MessageBox.Show(string.Format("{0} : {1} {2}",
                                              WordToTFS.Properties.Resources.RequiredFields_Fields_plural,
                                              requiredFields.Remove(requiredFields.Length - 2, 2),
                                              WordToTFS.Properties.Resources.RequiredFields_Required_plural));
                return;
            }
            IsCreated = true;
            Close();
        }

        public bool IsInteger(string input)
        {
            try
            {
                int.Parse(input);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        } 
        private void CancelButton1Click(object sender, RoutedEventArgs e)
        {
            IsCancelled = true;
            Close();
        }

        public void DataWindowClosed(object sender, CancelEventArgs e)
        {
            IsCancelled = true;
        }
    }
}
