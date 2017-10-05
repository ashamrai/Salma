using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using WordToTFSWordAddIn.Views;
using System.Linq;
using System.Threading;
using System.Windows;
using System.ServiceModel.Channels;
using MessageBox = System.Windows.MessageBox;

namespace WordToTFS
{
    internal class QueryManager
    {
        public enum QueryTypes
        {
            Folder,
            MyQ,
            TeamQ,
            FView,
            DView,
            HView,
            None
        }

        private readonly Project project;
        public TreeView Tree { set; get; }
        public WorkItemStore ItemStore { set; get; }

        public QueryManager(Project project, TreeView tree, WorkItemStore itemStore)
        {
            ItemStore = itemStore;
            this.project = project;
            Tree = tree;
            project.QueryHierarchy.Refresh();
            BuildQueryHierarchy(project.QueryHierarchy);
        }

        public event Action<QueryDefinition> Query;

        /// <summary> 
        /// Getting Project Query Hierarchy and define the the folders under it. 
        /// </summary> 
        /// <param name="queryHierarchy"></param> 
        void BuildQueryHierarchy(IEnumerable<QueryItem> queryHierarchy)
        {
            var root = new TreeViewItem { Header = project.Name };

            foreach (QueryFolder query in queryHierarchy)
            {
                DefineFolder(query, root);
            }

            Tree.Items.Add(root);
        }

        /// <summary>
        /// Define Query Folders under TreeView
        /// </summary>
        /// <param name="query">Query Folder</param>
        /// <param name="father"></param>
        void DefineFolder(QueryFolder query, TreeViewItem father)
        {
            father.IsExpanded = true;
            var item = new TreeViewItem { IsExpanded = true };

            var type = QueryTypes.Folder;

            if (query.IsPersonal)
            {
                type = QueryTypes.MyQ;
            }
            else if (query.Name == "Team Queries")
            {
                type = QueryTypes.TeamQ;
            }

            item.Header = CreateTreeItem(query.Name, type);
            father.Items.Add(item);

            foreach (var subQuery in query)
            {
                if (subQuery.GetType() == typeof(QueryFolder))
                {
                    DefineFolder((QueryFolder)subQuery, item);
                }
                else
                {
                    DefineQuery((QueryDefinition)subQuery, item);
                }
            }
        }

        /// <summary>
        /// Add Query Definition under a specific Query Folder.
        /// </summary>
        /// <param name="query">Query Definition - Contains the Query Details</param>
        /// <param name="queryFolder">Parent Folder</param>
        void DefineQuery(QueryDefinition queryDef, TreeViewItem queryFolder)
        {
            var item = new TreeViewItem();
            QueryTypes type;

            switch (queryDef.QueryType)
            {
                case QueryType.List: type = QueryTypes.FView; break;
                case QueryType.OneHop: type = QueryTypes.DView; break;
                case QueryType.Tree: type = QueryTypes.HView; break;
                default: type = QueryTypes.None; break;
            }
            item.Header = CreateTreeItem(queryDef.Name + " [...]", type);
            item.Tag = queryDef.Id;
            item.Selected += ItemOnSelected;
            item.Unselected += ItemUnselected;
            queryFolder.Items.Add(item);
        }


        public void ItemUnselected(object sender, RoutedEventArgs e)
        {
            if (Query != null)
            {

                Query(null);

            }

        }
  

        public void ItemOnSelected(object sender, RoutedEventArgs e)
        {
            try
            {
                var item = (TreeViewItem)sender;
                if (Query != null)
                {
                    var queryDef = (QueryDefinition)project.QueryHierarchy.Find((Guid)item.Tag);

                    queryDef.QueryText = queryDef.QueryText.Replace("@project", "'" + project.Name + "'");
                    Query(queryDef);


                    var query = new Query(ItemStore, queryDef.QueryText);
                    Int32 count = 0;
                    if (query.IsLinkQuery)
                    {
                        var queryResults = query.RunLinkQuery();
                        count = queryResults.Count();
                    }
                    else
                    {
                        var queryResults = query.RunQuery();
                        count = queryResults.Count;
                    }


                    QueryTypes type;

                    switch (queryDef.QueryType)
                    {
                        case QueryType.List: type = QueryTypes.FView; break;
                        case QueryType.OneHop: type = QueryTypes.DView; break;
                        case QueryType.Tree: type = QueryTypes.HView; break;
                        default: type = QueryTypes.None; break;
                    }
                    item.Header = CreateTreeItem(queryDef.Name + " [" + count + "]", type);
                    QueryReport.correctQuery = true;
                }
            }
            catch (InvalidQueryTextException ex)
            {
                MessageBox.Show(ex.Message, ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButton.OK, MessageBoxImage.Error);
                QueryReport.correctQuery = false;
            }
        }

        private StackPanel CreateTreeItem(string value, QueryTypes type)
        {
            var stake = new StackPanel { Orientation = Orientation.Horizontal };

            var img = new Image { Stretch = Stretch.Uniform, Source = GetImage(type) };
            var lbl = new Label { Content = value };

            stake.Children.Add(img);
            stake.Children.Add(lbl);

            return stake;
        }

        static ImageSource GetImage(QueryTypes type)
        {
            try
            {
                switch (type)
                {
                    case QueryTypes.MyQ:
                        return WindowFactory.IconConverter(Icons.MyQueries, 16, 16);
                    case QueryTypes.TeamQ:
                        return WindowFactory.IconConverter(Icons.TeamQueries, 16, 16);
                    case QueryTypes.Folder:
                        return WindowFactory.IconConverter(Icons.Folder, 16, 16);
                    case QueryTypes.FView:
                        return WindowFactory.IconConverter(Icons.FlatView, 16, 16);
                    case QueryTypes.DView:
                        return WindowFactory.IconConverter(Icons.DirectView, 16, 16);
                    case QueryTypes.HView:
                        return WindowFactory.IconConverter(Icons.HierarchicalView, 16, 16);
                    default:
                        return null;
                }
            }
            catch 
            {
                return null;
            }
        }
    }
}
