using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using WordToTFSWordAddIn.Views;
using Microsoft.Office.Interop.Word;
using MessageBox = System.Windows.MessageBox;

namespace WordToTFS
{
    using Microsoft.TeamFoundation.WorkItemTracking.Client;
    using System.Text.RegularExpressions;


    public class LinkWorkItem
    {
        private readonly TfsManager tfsManager;
        private Document document;
        
        public LinkWorkItem(Document activeDocument)
        {
            tfsManager = TfsManager.Instance;
            document = activeDocument;
        }
        /// <summary>
        /// The link item.
        /// </summary>
        /// <param name="itemsSource">
        /// The items source.
        /// </param>
        /// <param name="wiId">
        /// The wi id.
        /// </param>
        /// <param name="project">
        /// The project.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool LinkItem(List<string> itemsSource, int wiId, string project)
        {
            bool IsLinked = false;
            WorkItem selectedWorkItem = tfsManager.GetWorkItem(wiId);

            if (selectedWorkItem != null)
            {
                List<string> links = tfsManager.GetAllWorkItemLinksTypes().Select(item => item.ForwardEnd.Name).ToList();
                links.AddRange(tfsManager.GetAllWorkItemLinksTypes().Select(item => item.ReverseEnd.Name));

                links = links.Distinct().ToList();
                links.Sort();
                var popup = new LinkItem { linkTypesComboBox = { ItemsSource = links, SelectedIndex = 0 }, linkTypesComboBoxOutside = { ItemsSource = links, SelectedIndex = 0 }, workItemsListBox = { ItemsSource = itemsSource }, Manager = tfsManager };
                List<string> types = tfsManager.GetWorkItemsTypeForCurrentProject(project);
                types.Add(ResourceHelper.GetResourceString("ALL"));
                types.Sort();
                popup.wiTypesComboBox.ItemsSource = types;
                popup.wiTypesComboBox.SelectedValue = ResourceHelper.GetResourceString("ALL");

                popup.wiTypesComboBox.SelectionChanged += (s, a) => { 
                    popup.workItemsListBox.ItemsSource = ((string)a.AddedItems[0] == ResourceHelper.GetResourceString("ALL")) ? 
                        itemsSource : 
                        itemsSource.Where(t => t.Contains(a.AddedItems[0].ToString())).ToList(); };

                //foreach (ListBoxItem item in popup.workItemsListBox.Items.Cast<ListBoxItem>())
                //{
                //    bool isLinked = IsItemLinked(selectedWorkItem, item.Content.ToString());
                //}

                popup.Create(null, Icons.LinkItems);

                if (!popup.IsCanceled)
                {
                    if (popup.existingDocumentTabItem.IsSelected)
                    {
                        IList workItems = popup.workItemsListBox.SelectedItems;
                            foreach (string item in workItems)
                            {
                                try
                                {
                                    //int id = TfsManager.ParseId(item);

                                    Match match = Regex.Match(item, @"\d+");
                                    int id = 0;

                                    if (match.Success && Int32.TryParse(match.Value, out id))
                                        LinkWorkItems(selectedWorkItem, id, popup.linkTypesComboBox.SelectedValue.ToString(), popup.commentTextBox.Text);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                    }

                    if (popup.outsideDocumentTabItem.IsSelected)
                    {
                            foreach (WorkItem item in popup.WorkItemsToLink)
                            {
                                try
                                {
                                    LinkWorkItems(selectedWorkItem, item.Id, popup.linkTypesComboBoxOutside.SelectedValue.ToString(), popup.commentTextBox.Text);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                    }

                    if (popup.hyperLinkTabItem.IsSelected)
                    {
                        if (!string.IsNullOrWhiteSpace(popup.hyperlinkTextBox.Text) && popup.hyperlinkTextBox.Text != "http://")
                        {
                            try
                            {
                                var wiHyperLink = new Hyperlink(popup.hyperlinkTextBox.Text) { Comment = popup.commentTextBox.Text };
                                selectedWorkItem.Links.Add(wiHyperLink);
                                selectedWorkItem.Save();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }

                    IsLinked = true;
                }
            }
            else
            {
                MessageBox.Show(string.Format(ResourceHelper.GetResourceString("MSG_ITEM_IS_NOT_EXISTS_IN_TFS"), wiId), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return IsLinked;
        }

        /// <summary>
        /// The link work items.
        /// </summary>
        /// <param name="sourceId">
        /// The source id.
        /// </param>
        /// <param name="targetIds">
        /// The target ids.
        /// </param>
        /// <param name="linkTypeEnd">
        /// The link type end.
        /// </param>
        public void LinkWorkItems(int sourceId, List<int> targetIds, string linkTypeEnd)
        {
            WorkItem wi = tfsManager.GetWorkItem(sourceId);
            foreach (int id in targetIds)
            {
                var wiLink = new WorkItemLink(tfsManager.ItemsStore.WorkItemLinkTypes["Is Produced By"].ReverseEnd, sourceId, id);
                wi.Links.Add(wiLink);
            }

            wi.Save();
        }

        /// <summary>
        /// Link Work Items in TFS
        /// </summary>
        /// <param name="source">
        /// The source.
        /// </param>
        /// <param name="targetWorkItemId">
        /// The target Work Item Id.
        /// </param>
        /// <param name="linkTypeEndName">
        /// The link Type End Name.
        /// </param>
        /// <param name="comments">
        /// The comments.
        /// </param>
        public void LinkWorkItems(WorkItem source, int targetWorkItemId, string linkTypeEndName, string comments)
        {
            WorkItem wItem = tfsManager.GetWorkItem(targetWorkItemId);
            if (wItem != null)
            {
                WorkItemLinkTypeEnd linkTypeEnd = tfsManager.GetAllWorkItemLinksTypes().LinkTypeEnds[linkTypeEndName];
                var link = new RelatedLink(linkTypeEnd, targetWorkItemId) { Comment = comments };
                source.Links.Add(link);
                source.Save();
            }
        }

        /// <summary>
        /// The is item linked.
        /// </summary>
        /// <param name="sourceItem">
        /// The source item.
        /// </param>
        /// <param name="item">
        /// The item.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool IsItemLinked(WorkItem sourceItem, string item)
        {
            int id = TfsManager.ParseId(item);
            WorkItem targetWorkItem = tfsManager.GetWorkItem(id);
            return sourceItem.Links.Cast<Link>().Any(l => targetWorkItem.Links.Contains(l));
        }
    }
}
