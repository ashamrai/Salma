using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using WordToTFS.Properties;
using WordToTFS.View;
using WordToTFS.ViewModel;
using WordToTFS.ViewModel.RequiredFields;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Windows.Data;
using System.Globalization;
using System.Threading;


namespace WordToTFS.Model
{
    public class CreateNewWi
    {
        /// <summary>
        /// The add work item for current project.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <param name="title">
        /// The title.
        /// </param>
        /// <param name="workItemTypeString">
        /// The work item type string.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>

        public static int AddWorkItemForCurrentProject(string projectName, string title, string workItemTypeString, string areapath = "", string linkend = "", int linkid = 0, string DocUrl = "")
        {
            WorkItemType workItemType =
                TfsManager.Instance.ItemsStore.Projects[projectName].WorkItemTypes[workItemTypeString];
            var wi = new WorkItem(workItemType) { Title = title };

            if (areapath != "" && areapath != projectName) wi.AreaPath = areapath;

            var requiredFields = GetRequiredFieldsForWorkItem(wi);

            var popup = new RequiredFields { DataContext = requiredFields };

            if (requiredFields.Count != 0)
            {
                //Thread.CurrentThread.CurrentCulture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
                //Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd.MM.yyyy";
                popup.Create(null, Icons.AddNewWorkItem);
                if (popup.IsCancelled && !popup.IsCreated)
                    return 0;
            }

            if (!SetRequiredFieldsForWorkIetm(requiredFields, wi))
                return 0;

            if (linkend != "" && linkid > 0)
                if (TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds.Contains(linkend))
                {
                    WorkItemLinkTypeEnd linkTypeEnd = TfsManager.Instance.ItemsStore.WorkItemLinkTypes.LinkTypeEnds[linkend];
                    wi.Links.Add(new RelatedLink(linkTypeEnd, linkid));
                }

            if (DocUrl != "") wi.Links.Add(new Hyperlink(DocUrl));

            var _save_errors = wi.Validate();

            if (_save_errors.Count > 0)
            {
                return 0;
            }

            wi.Save();
            return wi.Id;
        }

        private static List<IViewModel> GetRequiredFieldsForWorkItem(WorkItem wi)
        {
            var requiredFields = new List<IViewModel>();
            foreach (Field field in wi.Fields)
            {
                if (field.IsRequired && field.Status != FieldStatus.Valid)
                {
                    if (field.FieldDefinition.FieldType == FieldType.String ||
                       field.FieldDefinition.FieldType == FieldType.PlainText ||
                        field.FieldDefinition.FieldType == FieldType.History)
                    {
                        if (field.AllowedValues.Count > 1)
                        {
                            List<string> comboboxFields = field.AllowedValues.Cast<string>().ToList();
                            bool isLimitedToAllowedValues = !field.IsLimitedToAllowedValues;
                            requiredFields.Add(new DropDownFieldViewModel(field.FieldDefinition.Name, comboboxFields, isLimitedToAllowedValues, false));

                        }
                        else
                        {
                            requiredFields.Add(new TextFieldViewModel(field.FieldDefinition.Name, false, ""));
                        }

                    }
                    else if (field.FieldDefinition.FieldType == FieldType.DateTime)
                    {
                        requiredFields.Add(new DateTimeFieldViewModel(field.FieldDefinition.Name));
                    }
                    else if (field.FieldDefinition.FieldType == FieldType.Html)
                    {
                        requiredFields.Add(new HtmlFieldViewModel(field.FieldDefinition.Name));
                    }
                    else if (field.FieldDefinition.FieldType == FieldType.Double ||
                             field.FieldDefinition.FieldType == FieldType.Integer)
                    {
                        if (field.AllowedValues.Count > 1)
                        {
                            List<string> comboboxFields = field.AllowedValues.Cast<string>().ToList();
                            bool isLimitedToAllowedValues = !field.IsLimitedToAllowedValues;
                            requiredFields.Add(new DropDownFieldViewModel(field.FieldDefinition.Name, comboboxFields, isLimitedToAllowedValues, true));
                        }
                        else
                        {
                            requiredFields.Add(new TextFieldViewModel(field.FieldDefinition.Name, true, "0"));
                        }

                    }
                    else if (field.FieldDefinition.FieldType == FieldType.Boolean)
                    {
                        requiredFields.Add(new BoolFieldViewModel(field.FieldDefinition.Name, new List<bool> { false, true }));
                    }
                }
            }
            return requiredFields;
        }

        private static bool SetRequiredFieldsForWorkIetm(List<IViewModel> requiredFields, WorkItem wi)
        {
            foreach (Field f in wi.Fields)
            {
                if (f.IsRequired && f.Status != FieldStatus.Valid)
                {
                    if (f.FieldDefinition.FieldType == FieldType.String || f.FieldDefinition.FieldType == FieldType.Html ||
                        f.FieldDefinition.FieldType == FieldType.PlainText ||
                        f.FieldDefinition.FieldType == FieldType.History)
                    {
                        foreach (var item in requiredFields)
                        {
                            if (f.FieldDefinition.Name == item.GetName())
                                f.Value = item.GetValue();
                        }

                    }
                    else if (f.FieldDefinition.FieldType == FieldType.DateTime)
                    {
                        foreach (var item in requiredFields)
                        {
                            if (f.FieldDefinition.Name == item.GetName())
                                f.Value = Convert.ToDateTime(item.GetValue());
                        }
                    }
                    else if (f.FieldDefinition.FieldType == FieldType.Double ||
                             f.FieldDefinition.FieldType == FieldType.Integer)
                    {
                        foreach (var item in requiredFields)
                        {
                            if (f.FieldDefinition.Name == item.GetName())
                                f.Value = Convert.ToDouble(item.GetValue());
                        }
                    }
                    else if (f.FieldDefinition.FieldType == FieldType.Boolean)
                    {
                        foreach (var item in requiredFields)
                        {
                            if (f.FieldDefinition.Name == item.GetName())
                                f.Value = Convert.ToBoolean(item.GetValue());
                        }
                    }

                    if (f.Status != FieldStatus.Valid)
                    {
                        MessageBox.Show(
                            string.Format(Resources.ERROR_MESSAGE_UNABLE_TO_CREATE_ITEM_WITH_REQUIRED_FIELD, f.Name),
                            ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
