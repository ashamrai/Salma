using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using WordToTFS;
using WordToTFS.ConfigHelpers;

namespace Salma2010
{
    public partial class WorkItemListControl : UserControl
    {
        public ThisAddIn wordaddin { get;  set; }
        public delegate void AddWIInfo(string pWiId, string pFieldName, string pTitle, string pControlId);
        public delegate void UpdateControlID(string pOldID, string pNewID);
        public AddWIInfo AddWiDelegate;
        public UpdateControlID UpdateControlIDDelegate;
        public String DocumentPath { get; set; }

        private class TagInfo
        {
            public string ControlId = "";
            public string WiId = "";
            public string WiField = "";
        }

        public WorkItemListControl()
        {
            InitializeComponent();
            AddWiDelegate = new AddWIInfo(AddWIInfoMethod);
            UpdateControlIDDelegate = new UpdateControlID(UpdateControlIDMethod);          
            
        }

        private void UpdateControlIDMethod(string pOldID, string pNewID)
        {
            foreach (TreeNode _tmpNode in treeWIList.Nodes)
            {
                foreach(TreeNode _fldNode in _tmpNode.Nodes)
                {
                    TagInfo _tg = (TagInfo)_fldNode.Tag;
                    if (_tg.ControlId == pOldID)
                    {
                        _tg.ControlId = pNewID;
                        return;
                    }
                }
            }
        }

        public void SyncWorkItems()
        {
            treeWIList.Nodes.Clear();

            List<ContentControl> controlsToUpdate = (from control in wordaddin.Application.ActiveDocument.ContentControls.OfType<ContentControl>()
                                                     select control).ToList<ContentControl>();

            if (controlsToUpdate.Count == 0) return;
            
            foreach(ContentControl _ctrl in controlsToUpdate)
            {
                string wiID = "", wiField = "", wiTitle = "";
                try { wiID = wordaddin.Application.ActiveDocument.Variables[_ctrl.ID + "_wiid"].Value; }
                catch { }
                try { wiField = wordaddin.Application.ActiveDocument.Variables[_ctrl.ID + "_wifieldEx"].Value; }
                catch { }
                try { wiTitle = wordaddin.Application.ActiveDocument.Variables[_ctrl.ID + "_wititle"].Value; }
                catch { }

                if (wiID != "" && wiField != "")
                    AddWIInfoMethod(wiID, wiField, wiTitle, _ctrl.ID);
            }
        }

        public void AddWIInfoMethod(string pWiId, string pFieldName, string pTitle, string pControlId)
        {
            TreeNode _winode = null;
            TreeNode _fieldnode = null;

            if (treeWIList.Nodes.Count > 0)
            {
                foreach (TreeNode _tmpnode in treeWIList.Nodes)
                {
                    TagInfo _tg = (TagInfo)_tmpnode.Tag;
                    if (_tg.WiId == pWiId)
                        _winode = _tmpnode;
                }
            }

            if (_winode == null)
            {
                _winode = new TreeNode();
                TagInfo _tg = new TagInfo();
                _tg.ControlId = pControlId;
                _tg.WiId = pWiId;
                _winode.Tag = _tg;
                treeWIList.Nodes.Add(_winode);
            }

            if (pTitle != "")
                _winode.Text = pWiId + ": " + ((pTitle.Length <= 70) ? pTitle : ( pTitle.Substring(0,67) + "..."));

            if (_winode.Nodes.Count > 0)
            {
                foreach (TreeNode _tmpnode in _winode.Nodes)
                {
                    TagInfo _tg = (TagInfo)_tmpnode.Tag;
                    if (_tg.WiField.ToString() == pFieldName)
                        _fieldnode = _tmpnode;
                }
            }

            if (_fieldnode == null)
            {
                _fieldnode = new TreeNode();
                TagInfo _tg = new TagInfo();
                _tg.ControlId = pControlId;
                _tg.WiId = pWiId;
                _tg.WiField = pFieldName;
                _fieldnode.Tag = _tg;
                _winode.Nodes.Add(_fieldnode);                
            }

            _fieldnode.Text = pFieldName;
        }

        private void treeWIList_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                TreeNode _node = treeWIList.SelectedNode;

                if (_node.Parent != null)
                {
                    TagInfo _tg = (TagInfo)_node.Tag;
                    wordaddin.Application.ActiveDocument.ContentControls[_tg.ControlId].Range.Select();
                }
            }
            catch(Exception)
            {

            }
        }

        private void btnExpandAll_Click(object sender, EventArgs e)
        {
            treeWIList.ExpandAll();           
        }

        private void btnCollapseAll_Click(object sender, EventArgs e)
        {
            treeWIList.CollapseAll();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            SyncWorkItems();
        }
    }
}
