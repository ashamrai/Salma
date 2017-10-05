using System.Windows;
using System.Windows.Controls;
using WordToTFS.ViewModel;
using WordToTFS.ViewModel.RequiredFields;

namespace WordToTFS.Helpers
{
    public class DataGridCellDataTemplateSelector : DataTemplateSelector
    {
        public DataTemplate StringTemplate { get; set; }
        public DataTemplate EditableComboBoxTemplate { get; set; }
        public DataTemplate NonEditableComboboxTemplate { get; set; }
        public DataTemplate HtmlTemplate { get; set; }
        public DataTemplate DateTimeTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is DateTimeFieldViewModel)
                return DateTimeTemplate;
            if (item is TextFieldViewModel)
                return StringTemplate;
            if (item is DropDownFieldViewModel)
            {
                var temp = (DropDownFieldViewModel)item;
                if (temp.IsEditable)
                    return EditableComboBoxTemplate;
                return NonEditableComboboxTemplate;
             }
            if (item is BoolFieldViewModel)
                return NonEditableComboboxTemplate;
            if (item is HtmlFieldViewModel)
                return HtmlTemplate;
            return base.SelectTemplate(item, container);
        }
    }
}