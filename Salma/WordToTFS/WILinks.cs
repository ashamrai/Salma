using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToTFS
{
    public class WILink
    {
        public string LinkTypeEnd {get; set;}
        public int TargetId { get; set; }
        public string TargetWorkItemType { get; set; }
        public string TargetWorkItemTitle { get; set; }
        public bool IsSelected { get; set; }
        public List<string> ActionTypes { get; set; }
        public string SelectedActionType { get; set; }

        public WILink()
        {
            ActionTypes = new List<string>();

            ActionTypes.Add(Properties.Resources.Obolete_List_Item_Copy);
            ActionTypes.Add(Properties.Resources.Obolete_List_Item_Move);
            SelectedActionType = Properties.Resources.Obolete_List_Item_Copy;

            LinkTypeEnd = "";
            TargetId = 0;
            TargetWorkItemTitle = "";
            TargetWorkItemType = "";
            IsSelected = true;
        }

        public void RemoveCopyAction()
        {
            ActionTypes.Remove(Properties.Resources.Obolete_List_Item_Copy);
            SelectedActionType = Properties.Resources.Obolete_List_Item_Move;
        }
    }
}
