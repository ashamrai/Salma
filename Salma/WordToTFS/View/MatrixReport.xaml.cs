using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using WordToTFS;
using System;

namespace WordToTFSWordAddIn.Views
{

    /// <summary>
    ///     Interaction logic for MatrixReport.xaml
    /// </summary>
    public partial class MatrixReport : Window
    {
        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MatrixReport"/> class.
        /// </summary>
        public MatrixReport()
        {
            this.InitializeComponent();
            this.IsCanceled = true;
            dateFrom.DisplayDateEnd = DateTime.Now;
            dateTo.DisplayDateEnd = DateTime.Now;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets a value indicating whether is canceled.
        /// </summary>
        public bool IsCanceled { get; set; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The get selected states.
        /// </summary>
        /// <param name="box">
        /// The box.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<string> GetSelectedStates(ComboBox box)
        {
            var stateSelectedItems = new List<string>();
            foreach (CheckBox chk in box.Items)
            {
                if (chk.IsChecked.Value && chk.Content.ToString() != ResourceHelper.GetResourceString("ALL"))
                {
                    stateSelectedItems.Add(chk.Content.ToString());
                }
            }

            return stateSelectedItems;
        }

        /// <summary>
        /// The init state values.
        /// </summary>
        /// <param name="combo">
        /// The combo.
        /// </param>
        /// <param name="states">
        /// The states.
        /// </param>
        public void InitStateValues(ComboBox combo, List<string> states)
        {
            List<string> st = states;
            st.Sort();
            st.Insert(0, ResourceHelper.GetResourceString("ALL"));
            combo.Items.Clear();

            ////TODO: add check boxes
            foreach (string item in st)
            {
                var cbx = new CheckBox();
                cbx.Width = 155;
                cbx.Content = item;
                cbx.IsChecked = true;
                cbx.Click += delegate(object sender, RoutedEventArgs e)
                    {
                        var selectedCbx = sender as CheckBox;
                        var checkedCount = combo.Items.Cast<CheckBox>().Count(c => c.IsChecked.Value);
                        if (selectedCbx.Content.ToString() == ResourceHelper.GetResourceString("ALL"))
                        {
                            this.SelectAll(combo, selectedCbx);
                        }

                        else if (checkedCount == 0)
                        {
                            combo.Text = ResourceHelper.GetResourceString("NOT_SELECTED");
                        }
                        else if (checkedCount > 1)
                        {
                            if (checkedCount >= combo.Items.Count - 1)
                            {
                                var chkSelectAll = (CheckBox)combo.Items[0];
                                chkSelectAll.IsChecked = true;
                                combo.Text = chkSelectAll.Content.ToString();
                            }
                            else
                            {
                                combo.Text = ResourceHelper.GetResourceString("CUSTOM");
                            }
                        }
                        else if (checkedCount == 1)
                        {
                            combo.Text = combo.Items.Cast<CheckBox>().FirstOrDefault(c => c.IsChecked.Value).Content.ToString();
                        }
                    };
                cbx.Checked += delegate(object sender, RoutedEventArgs e)
                    {
                        var selectedCbx = sender as CheckBox;
                        if (selectedCbx.Content.ToString() != ResourceHelper.GetResourceString("ALL"))
                        {
                            var chkSelectAll = (CheckBox)combo.Items[0];
                            if (combo.Items.Cast<CheckBox>().Count(c => c.IsChecked.Value) >= combo.Items.Count-1)
                            {
                                chkSelectAll.IsChecked = true;
                                combo.Text = chkSelectAll.Content.ToString();
                            }
                        }
                    };

                cbx.Unchecked += delegate(object sender, RoutedEventArgs e)
                    {
                        var selectedCbx = sender as CheckBox;
                        if (selectedCbx.Content.ToString() != ResourceHelper.GetResourceString("ALL"))
                        {
                            var chkSelectAll = (CheckBox)combo.Items[0];
                            chkSelectAll.IsChecked = false;
                            combo.Text = ResourceHelper.GetResourceString("CUSTOM");
                        }
                    };

                combo.Items.Add(cbx);
            }

            combo.SelectedIndex = 0;
        }

        #endregion

        #region Methods

        /// <summary>
        /// The cancel button click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            this.IsCanceled = true;
            this.Close();
        }

        /// <summary>
        /// The date picker_ loaded.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void DatePicker_Loaded(object sender, RoutedEventArgs e)
        {
            var picker = sender as DatePicker;
            FieldInfo fiTextBox = picker.GetType().GetField("_textBox", BindingFlags.Instance | BindingFlags.NonPublic);

            if (fiTextBox != null)
            {
                var dateTextBox = (DatePickerTextBox)fiTextBox.GetValue(picker);

                if (dateTextBox != null)
                {
                    PropertyInfo piWatermark = dateTextBox.GetType().GetProperty("Watermark", BindingFlags.Instance | BindingFlags.NonPublic);

                    if (piWatermark != null)
                    {
                        piWatermark.SetValue(dateTextBox, WordToTFS.Properties.Resources.DatePickerWatermark, null);
                    }
                }
            }
        }

        /// <summary>
        /// The insert button click.
        /// </summary>
        /// <param name="sender">
        /// The sender.
        /// </param>
        /// <param name="e">
        /// The e.
        /// </param>
        private void InsertButtonClick(object sender, RoutedEventArgs e)
        {
            if (this.dateFrom.SelectedDate != null && this.dateTo.SelectedDate != null && this.dateFrom.SelectedDate > this.dateTo.SelectedDate)
            {
                MessageBox.Show(ResourceHelper.GetResourceString("MSG_ERROR_DATE"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error); 
            }
            else
            {
                this.IsCanceled = false;
                this.Close();
            }
        }

        /// <summary>
        /// The select all.
        /// </summary>
        /// <param name="combo">
        /// The combo.
        /// </param>
        /// <param name="cbxItem">
        /// The cbx item.
        /// </param>
        private void SelectAll(ComboBox combo, CheckBox cbxItem)
        {
            foreach (CheckBox lbxItem in combo.Items)
            {
                lbxItem.IsChecked = cbxItem.IsChecked;
            }

            combo.Text = cbxItem.IsChecked.Value ? cbxItem.Content.ToString() : ResourceHelper.GetResourceString("NOT_SELECTED");
        }

        #endregion
    }
}