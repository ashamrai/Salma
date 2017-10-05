
namespace WordToTFS
{
    using System;
    using System.Diagnostics;
    using System.Timers;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows;

    using WordToTFS.Model;
    using WordToTFS.View;

    using WordToTFSWordAddIn.Views;

    using Microsoft.Office.Interop.Word;
    using Microsoft.TeamFoundation.WorkItemTracking.Client;

    using Timer = System.Timers.Timer;

    /// <summary>
    /// Traceability Matrix
    /// </summary>
    public class TraceabilityMatrix
    {
        // TODO: Continue refactor
        #region Fields

        /// <summary>
        /// Document.
        /// </summary>
        private Document document;

        /// <summary>
        /// Tfs manager.
        /// </summary>
        private readonly TfsManager tfsManager;

        /// <summary>
        /// The cancellation token source.
        /// </summary>
        private CancellationTokenSource cancellationTokenSource;

        /// <summary>
        /// Matrix Report Dialog Data
        /// </summary>
        private MatrixReportData matrixReportData = null;

        /// <summary>
        /// Progress dialog.
        /// </summary>
        private ProgressDialog progressDialog;
      

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TraceabilityMatrix"/> class.
        ///     Traceability Matrix Constructor
        /// </summary>
        /// <param name="manager">
        /// The manager.
        /// </param>
        /// <param name="activeDocument">
        /// The active Document.
        /// </param>
        public TraceabilityMatrix(Document activeDocument)
        {
            tfsManager = TfsManager.Instance;
            document = activeDocument;
            matrixReportData = new MatrixReportData();
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Generate Traceability Matrix Report
        /// </summary>
        /// <param name="project">
        /// Project Name
        /// </param>
        public void GenerateMatrix(string project)
        {
            var popup = new MatrixReport();
            InitMatrixPopUp(popup, project);

            popup.Create(null, Icons.TraceabilityMatrix);

            if (!popup.IsCanceled)
            {
                matrixReportData.ColumnsTypesSelectedItem = popup.HorizontalTypes.SelectedItem.ToString();
                matrixReportData.RowsTypesSelectedItem = popup.VerticalTypes.SelectedItem.ToString();
                matrixReportData.RelatedSelectedItem = popup.Relateds.SelectedItem.ToString();

                if (popup.chkIncludeNotLinked.IsChecked != null)
                {
                    matrixReportData.IncludeNotLinkedItems = popup.chkIncludeNotLinked.IsChecked.Value;
                }
   
                matrixReportData.DateFrom = popup.dateFrom.SelectedDate;
                matrixReportData.DateTo = popup.dateTo.SelectedDate;
                matrixReportData.StateColumnsSelectedItem = popup.GetSelectedStates(popup.StateHorisontal);
                matrixReportData.StateRowsSelectedItem = popup.GetSelectedStates(popup.StateVertical);

                progressDialog = new ProgressDialog();
                string uiCulture = Thread.CurrentThread.CurrentUICulture.ToString();
                progressDialog.Execute(
                    (cancelTokenSource) =>
                        {
                            Thread.CurrentThread.CurrentUICulture = new CultureInfo(uiCulture);
                            cancellationTokenSource = cancelTokenSource;
                            var workItemsInColumn = new List<WorkItem>();
                            var workItemsInRow = new List<WorkItem>();
                            if (!cancellationTokenSource.IsCancellationRequested)
                            {
                                progressDialog.UpdateProgress(0, ResourceHelper.GetResourceString("MSG_RETRIEVING_DATA"), true);

                                MatrixData data = GetItemsByMatrixFilter(project);

                                workItemsInRow = data.Rows;
                                workItemsInColumn = data.Columns;

                                if (workItemsInRow.Count == 0 || workItemsInColumn.Count == 0)
                                {
                                    MessageBox.Show(ResourceHelper.GetResourceString("MSG_DATA_NOT_AVAILABLE"), string.Empty, MessageBoxButton.OK, MessageBoxImage.Information);
                                    return;
                                }

                                workItemsInRow.Sort((x, y) => x.Id.CompareTo(y.Id));
                                workItemsInColumn.Sort((x, y) => x.Id.CompareTo(y.Id));

                                DrawMatrixReport(project, workItemsInRow, workItemsInColumn);
                            }
                        });

                progressDialog.Create(string.Format(ResourceHelper.GetResourceString("MSG_TRACEBILITY_MATRIX"), project), Icons.TraceabilityMatrix);
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// The build query.
        /// </summary>
        /// <param name="projectName">
        /// The project name.
        /// </param>
        /// <param name="selectedItemType">
        /// The selected item type.
        /// </param>
        /// <param name="stateCondition">
        /// The state condition.
        /// </param>
        /// <param name="dateCondition">
        /// The date condition.
        /// </param>
        /// <param name="useWorkItemsTable">
        /// Use Work Items or WorkItemLinks Table.
        /// </param>
        /// <param name="relatedStateCondition">
        /// The related state condition.
        /// </param>
        /// <returns>
        /// query text
        /// </returns>
        private static string BuildQuery(string projectName, string selectedItemType, string stateCondition, string dateCondition, bool useWorkItemsTable, string relatedStateCondition = "")
        {
            string queryText;

            if (useWorkItemsTable)
            {
                queryText = string.Format(@"SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '{0}' AND ([Work Item Type] = '{1}') {2} {3}", projectName, selectedItemType, stateCondition, dateCondition);
                if(TfsManager.Instance.StateName == "Состояние")
                queryText = string.Format(@"SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '{0}' AND ([Тип рабочего элемента] = '{1}') {2} {3}", projectName, selectedItemType, stateCondition, dateCondition);
            }
            else
            {
                relatedStateCondition = relatedStateCondition == string.Empty ? relatedStateCondition : string.Format("AND {0}", relatedStateCondition);
                queryText = string.Format("SELECT [System.Id] FROM WorkItemLinks WHERE  (Source.[System.TeamProject] = '{0}') AND ([Source].[System.WorkItemType] = '{1}') {2} {3} {4}", projectName, selectedItemType, stateCondition, dateCondition, relatedStateCondition);
            }

            return queryText;
        }

        /// <summary>
        ///     Build Query Date Condition
        /// </summary>
        /// <returns>query text</returns>
        private string BuildQueryDateCondition()
        {
            string dateColumnName = matrixReportData.IncludeNotLinkedItems ? "[Created Date]" : "Source.[Created Date]";
            if(TfsManager.Instance.StateName == "Состояние")
                dateColumnName = matrixReportData.IncludeNotLinkedItems ? "[Дата создания]" : "Source.[Дата создания]";
            string dateCondition = string.Empty;

            if (matrixReportData.DateFrom != null && matrixReportData.DateTo != null)
            {
                dateCondition = string.Format("AND ({0} >= '{1}' AND {0} <= '{2}')", dateColumnName, matrixReportData.DateFrom.Value.ToShortDateString(), matrixReportData.DateTo.Value.ToShortDateString());
            }
            else
            {
                if (matrixReportData.DateFrom != null)
                {
                    dateCondition = string.Format("AND ({0} >= '{1}')", dateColumnName, matrixReportData.DateFrom.Value.ToShortDateString());
                }
                else if (matrixReportData.DateTo != null)
                {
                    dateCondition = string.Format("AND ({0} <= '{1}')", dateColumnName, matrixReportData.DateTo.Value.ToShortDateString());
                }
            }

            return dateCondition;
        }

        /// <summary>
        /// Build query state condition.
        /// </summary>
        /// <param name="stateSelectedItems">
        /// State selected items.
        /// </param>
        /// <returns>
        /// query text
        /// </returns>
        private string BuildQueryStateCondition(List<string> stateSelectedItems)
        {
            var stateBuilder = new StringBuilder();
            if (stateSelectedItems.Count > 0)
            {
                stateBuilder.AppendFormat("AND {0}[{1}] IN (", matrixReportData.IncludeNotLinkedItems ? string.Empty : "[Source].", TfsManager.Instance.StateName);
                //stateBuilder.AppendFormat("AND {0}[State] IN (", matrixReportData.IncludeNotLinkedItems ? string.Empty : "[Source].");
                foreach (string item in stateSelectedItems)
                {
                    stateBuilder.AppendFormat("'{0}',", item);
                }

                stateBuilder.Remove(stateBuilder.Length - 1, 1).Append(")");
            }

            return stateBuilder.ToString();
        }

        /// <summary>
        /// Draw Matrix Table
        /// </summary>
        /// <param name="project">
        /// The project.
        /// </param>
        /// <param name="workItemsInRow">
        /// Work Items In Row
        /// </param>
        /// <param name="workItemsInColumn">
        /// Work Items In Column
        /// </param>
        private void DrawMatrixReport(string project, List<WorkItem> workItemsInRow, List<WorkItem> workItemsInColumn)
        {
            SetMatrixTitle(project, ResourceHelper.GetResourceString("WORD_TRACEABILITY_MATRIX"));
            const int TABLE_COLUMNS_COUNT = 10;
            int pages = workItemsInColumn.Count / TABLE_COLUMNS_COUNT;
            if ((workItemsInColumn.Count % TABLE_COLUMNS_COUNT) > 0)
            {
                pages++;
            }

            for (int i = 0; i < pages; i++)
            {
                if (!cancellationTokenSource.IsCancellationRequested)
                {
                    progressDialog.UpdateProgress(Convert.ToInt32((i / (decimal)pages) * 100), string.Format(ResourceHelper.GetResourceString("MATRIX_PROGRESS"), i + 1, pages), true);
                    List<WorkItem> workItemsInColumnPerPage = workItemsInColumn.Skip(TABLE_COLUMNS_COUNT * i).Take(TABLE_COLUMNS_COUNT).ToList();
                    DrawTable(project, workItemsInRow, workItemsInColumnPerPage);
                }
            }
        }

        /// <summary>
        /// Build Data String Table
        /// </summary>
        /// <param name="workItemsInRow"></param>
        /// <param name="workItemsInColumn"></param>
        /// <returns></returns>
        private string BuildDataString(List<WorkItem> workItemsInRow, List<WorkItem> workItemsInColumn)
        {
            const int TITLE_ROW_INDEX = 1;
            const int HEADER_ROW_INDEX = 2;
            const int VERTICAL_STATUS_ROW_INDEX = 3;
            const int DESCRIPTION_ROW_INDEX = 4;
            const int PROJECT_COLUMN_INDEX = 1;
            const int HORIZONTAL_STATUS_COLUMN_INDEX = 2;
            const int ROW_SHIFT = 5;
            const int COL_SHIFT = 3;

            var dataString = new StringBuilder();
            var columnsCount = workItemsInColumn.Count + HEADER_ROW_INDEX; 
            var rowsCount = workItemsInRow.Count + DESCRIPTION_ROW_INDEX;

            for (var row = 1; row <= rowsCount; row++)
            {
                for (var col = 1; col <= columnsCount; col++)
                {
                    if (row == 1 && col == 1)
                    {
                        dataString.Append(" ");
                    }
                    else
                    {
                        // Row: populate item Id with hyperlink
                        if (row > DESCRIPTION_ROW_INDEX && col == PROJECT_COLUMN_INDEX)
                        {
                            dataString.Append(workItemsInRow[row - ROW_SHIFT].Id);
                        }

                        // Row: populate item State
                        if (col == HORIZONTAL_STATUS_COLUMN_INDEX && row > DESCRIPTION_ROW_INDEX)
                        {
                            dataString.Append(workItemsInRow[row - ROW_SHIFT].State);
                        }

                        // Row: populate item type
                        if (row == DESCRIPTION_ROW_INDEX && col == PROJECT_COLUMN_INDEX)
                        {
                            var wi = workItemsInRow.FirstOrDefault();
                            if (wi != null)
                            {
                                dataString.Append(wi.Type.Name);
                            }
                        }

                        // Column: populate item Type
                        if (col == PROJECT_COLUMN_INDEX && row == HEADER_ROW_INDEX)
                        {
                            var wi = workItemsInColumn.FirstOrDefault();
                            if (wi != null)
                            {
                                dataString.Append(wi.Type.Name);
                            }
                        }

                        // Column: populate item Title
                        if (col > HORIZONTAL_STATUS_COLUMN_INDEX && row == TITLE_ROW_INDEX)
                        {
                            dataString.Append(workItemsInColumn[col - COL_SHIFT].Title.Trim().Replace("\t", " ").Replace("\n", " "));
                        }

                        // Column: populate item Id with hyperlink
                        if (col > HORIZONTAL_STATUS_COLUMN_INDEX && row == HEADER_ROW_INDEX)
                        {
                            dataString.Append(workItemsInColumn[col - COL_SHIFT].Id);
                        }

                        // Column: populate item State
                        if (col > HORIZONTAL_STATUS_COLUMN_INDEX && row == VERTICAL_STATUS_ROW_INDEX)
                        {
                            dataString.Append(workItemsInColumn[col - COL_SHIFT].State);
                        }

                        // Mark as Linked
                        if (row >= ROW_SHIFT && col >= COL_SHIFT)
                        {
                            if (IsLinked(workItemsInRow[row - ROW_SHIFT], workItemsInColumn[col - COL_SHIFT]))
                            {
                                dataString.Append("✔");
                            }
                        }
                    }

                    // Append a field delimiter (\t). If we're on the last column, so append a record delimiter (\n).
                    dataString.Append(col < columnsCount ? "\t" : "\n");
                }
            }

            return dataString.ToString();
        } 

        /// <summary>
        /// Draw Table
        /// </summary>
        /// <param name="projectName">
        /// </param>
        /// <param name="workItemsInRow">
        /// </param>
        /// <param name="workItemsInColumn">
        /// </param>
        private void DrawTable(string projectName, List<WorkItem> workItemsInRow, List<WorkItem> workItemsInColumn)
        {
            const int TITLE_ROW_INDEX = 1;
            const int HEADER_ROW_INDEX = 2;
            const int VERTICAL_STATUS_ROW_INDEX = 3;
            const int DESCRIPTION_ROW_INDEX = 4;
            const int PROJECT_COLUMN_INDEX = 1;
            const int HORIZONTAL_STATUS_COLUMN_INDEX = 2;
            const int ROW_SHIFT = 5;
            const int COL_SHIFT = 3;

            // for additional info see here: "http://msdn.microsoft.com/en-us/library/office/aa537149(v=office.11).aspx"
            object objDefaultBehaviorWord8 = WdDefaultTableBehavior.wdWord8TableBehavior;
            object objAutoFitFixed = WdAutoFitBehavior.wdAutoFitFixed;

            var parRange = OfficeHelper.CreateParagraphRange(ref document);
            
            parRange.Text = BuildDataString(workItemsInRow, workItemsInColumn);
            try
            {
                Table table = parRange.ConvertToTable(AutoFitBehavior: objAutoFitFixed, DefaultTableBehavior: objDefaultBehaviorWord8);

                table.set_Style(WdBuiltinStyle.wdStyleTableLightGrid);

                table.Select();
                Selection sel = document.Application.Selection;
                if (sel != null)
                {
                    sel.Font.Bold = 0;
                    sel.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    sel.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // style titles

                sel = SelectCells(table, 1, 3, 1, workItemsInColumn.Count + 2); // 2 cells is a title for items in rows.

                if (sel != null)
                {
                    sel.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    sel.Font.Size = sel.Range.Font.Size - 1;
                    sel.Font.Bold = 1;
                    sel.Orientation = WdTextOrientation.wdTextOrientationUpward;
                    sel.Cells.Height = 100;
                    sel.EscapeKey();
                }
                // Set workItems types as bold
                sel = SelectCells(table, 1, 1, 4, 1);
                if (sel != null)
                {
                    sel.Font.Bold = 1;
                    sel.EscapeKey();
                }
                
                // Add hyperlinks to work items in column
                Parallel.For(0, workItemsInColumn.Count, i => table.Columns[i + COL_SHIFT].Cells[2].Range.Hyperlinks.Add(table.Columns[i + COL_SHIFT].Cells[2].Range, tfsManager.GetWorkItemLink(workItemsInColumn[i].Id, projectName)));

                // Add hyperlinks to work items in row
                Parallel.For(0, workItemsInRow.Count, i => table.Rows[i + ROW_SHIFT].Cells[1].Range.Hyperlinks.Add(table.Rows[i + ROW_SHIFT].Cells[1].Range, tfsManager.GetWorkItemLink(workItemsInRow[i].Id, projectName)));


                table.Rows[DESCRIPTION_ROW_INDEX].Cells[PROJECT_COLUMN_INDEX].Merge(table.Rows[DESCRIPTION_ROW_INDEX].Cells[PROJECT_COLUMN_INDEX + 1]);
                table.Rows[2].Cells[1].Merge(table.Rows[2].Cells[2]);
                table.Rows[1].Cells[1].Merge(table.Rows[1].Cells[2]);
                table.Rows[1].Cells[1].Merge(table.Rows[2].Cells[1]);
                table.Cell(3, 1).Merge(table.Cell(3, 2));
                table.Cell(3, 1).Merge(table.Cell(1, 1));
                //// Merge rows
                //table.Rows[4].Cells[1].Merge(table.Rows[4].Cells[2]);
                //sel = SelectCells(table, 1, 1, 3, 2);
                //if (sel != null)
                //{
                //    sel.Cells.Merge();
                //}

                document.Content.InsertParagraphAfter();
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// The select cells.
        /// </summary>
        /// <param name="tbl">
        /// The tbl.
        /// </param>
        /// <param name="rowStart">
        /// The row start.
        /// </param>
        /// <param name="colStart">
        /// The col start.
        /// </param>
        /// <param name="rowEnd">
        /// The row end.
        /// </param>
        /// <param name="colEnd">
        /// The col end.
        /// </param>
        /// <returns>
        /// The <see cref="Selection"/>.
        /// </returns>
        private Selection SelectCells(Table tbl, int rowStart, int colStart, int rowEnd, int colEnd)
        {
            object objWdCell = WdUnits.wdCell;
            object objWdCharacter = WdUnits.wdCharacter;
            object objWdLine = WdUnits.wdLine;

            Selection sel = null;
            int nrRows = tbl.Rows.Count;
            int nrCols = tbl.Columns.Count;
            // Make sure the start points exist in the table.
            // If they don't, then return without 
            // setting the selection
            if (rowStart > nrRows)
                return sel;
            if (colStart > nrCols)
                return sel;
            // Make sure the end point exists in the table.
            // If it does not, then set the last row/column as end points.
            if (rowEnd >= nrRows)
                rowEnd = (nrRows - rowStart + 1);
            if (colEnd >= nrCols)
                colEnd = (nrCols - colStart + 1);
            // Select the start cell.
            tbl.Cell(rowStart, colStart).Select();
            sel = document.Application.Selection;
            // Make sure the selection will extend.
            sel.ExtendMode = true;
            // First select the start cell.
            sel.Expand(ref objWdCell);
            // Next extend across the columns.
            // Subtract one from colEnd and rowEnd because the first row 
            // and column are already selected.
            object objColEnd = (object)(colEnd - 1);
            object objRowEnd = (object)(rowEnd - 1);
            sel.MoveRight(ref objWdCharacter, ref objColEnd, null);
            // Now extend down the rows.
            sel.MoveDown(ref objWdLine, ref objRowEnd, null);

            return sel;
        }

        /// <summary>
        /// Get Items By Selected Matrix Filter
        /// </summary>
        /// <param name="projectName">
        /// project Name
        /// </param>
        /// <returns>
        /// Work Item list
        /// </returns>
        private MatrixData GetItemsByMatrixFilter(string projectName)
        {
            var data = new MatrixData();

            string rowStateCondition = BuildQueryStateCondition(matrixReportData.StateRowsSelectedItem);
            string columnStateCondition = BuildQueryStateCondition(matrixReportData.StateColumnsSelectedItem);
            string relatedStateCondition = matrixReportData.RelatedSelectedItem == ResourceHelper.GetResourceString("ALL") ? string.Empty : string.Format("[System.Links.LinkType] = '{0}'", matrixReportData.RelatedSelectedItem);
            string dateCondition = BuildQueryDateCondition();

            string queryText;
            if (matrixReportData.IncludeNotLinkedItems)
            {
                // Rows
                queryText = BuildQuery(projectName, matrixReportData.RowsTypesSelectedItem, rowStateCondition, dateCondition, true, string.Empty);
                data.Rows = tfsManager.ExecuteQuery(queryText).Cast<WorkItem>().ToList();

                // Columns
                queryText = BuildQuery(projectName, matrixReportData.ColumnsTypesSelectedItem, columnStateCondition, dateCondition, true, string.Empty);
                data.Columns = tfsManager.ExecuteQuery(queryText).Cast<WorkItem>().ToList();
            }
            else
            {
                // columns
                queryText = BuildQuery(projectName, matrixReportData.ColumnsTypesSelectedItem, columnStateCondition, dateCondition, false);
                var q = new Query(tfsManager.ItemsStore, queryText);
                WorkItemLinkInfo[] colIds = q.RunLinkQuery();

                // rows
                queryText = BuildQuery(projectName, matrixReportData.RowsTypesSelectedItem, rowStateCondition, dateCondition, false, relatedStateCondition);
                q = new Query(tfsManager.ItemsStore, queryText);
                WorkItemLinkInfo[] rowIds = q.RunLinkQuery();

                foreach (WorkItemLinkInfo rowId in rowIds)
                {
                    foreach (WorkItemLinkInfo colId in colIds)
                    {
                        if (colId.TargetId == rowId.SourceId && colId.SourceId != 0)
                        {
                            var rowItem = tfsManager.GetWorkItem(rowId.SourceId);
                            var colItem = tfsManager.GetWorkItem(colId.SourceId);
                            if (IsLinked(rowItem, colItem))
                            {
                                if (data.Rows.All(t => t.Id != rowId.SourceId))
                                {
                                    data.Rows.Add(rowItem);
                                }

                                if (data.Columns.All(t => t.Id != colId.SourceId))
                                {
                                    data.Columns.Add(colItem);
                                }
                            }
                        }
                    }
                }
            }

            return data;
        }

        /// <summary>
        /// Init Matrix Report popup with default values.
        /// </summary>
        /// <param name="popup">
        /// MatrixReport
        /// </param>
        /// <param name="project">
        /// project name
        /// </param>
        private void InitMatrixPopUp(MatrixReport popup, string project)
        {
            Project proj = tfsManager.GetProject(project);
            List<WorkItemType> types = proj.WorkItemTypes.Cast<WorkItemType>().OrderBy(f => f.Name).ToList();
            IEnumerable<string> itemTypes = types.Where(f => f.Name != "Code Review Request" && f.Name != "Code Review Response" && f.Name != "Feedback Request" && f.Name != "Feedback Response" && f.Name != "Shared Steps").Select(f => f.Name);
            itemTypes = itemTypes.Where(f => f != "Запрос на проверку кода" && f != "Ответ на проверку кода" && f != "Ответ на отзыв" && f != "Запрос отзыва" && f != "Общие шаги").Select(f => f);
            popup.HorizontalTypes.SelectionChanged += (s, a) =>
                {
                    if (popup.HorizontalTypes.SelectedValue != null)
                    {
                        List<string> states = tfsManager.GetWorkItemStatesByType(project, popup.HorizontalTypes.SelectedValue.ToString());
                        popup.InitStateValues(popup.StateHorisontal, states);
                    }
                };
            popup.VerticalTypes.SelectionChanged += (s, a) =>
                {
                    if (popup.HorizontalTypes.SelectedValue != null)
                    {
                        List<string> states = tfsManager.GetWorkItemStatesByType(project, popup.VerticalTypes.SelectedValue.ToString());
                        popup.InitStateValues(popup.StateVertical, states);
                    }
                };
            popup.HorizontalTypes.ItemsSource = itemTypes;
            popup.HorizontalTypes.SelectedItem = 0;
            popup.VerticalTypes.ItemsSource = itemTypes;
            popup.VerticalTypes.SelectedIndex = 0;

            List<string> wiLinkTypes = proj.Store.WorkItemLinkTypes.Select(f => f.ForwardEnd.Name).ToList();
            wiLinkTypes.AddRange(proj.Store.WorkItemLinkTypes.Select(f => f.ReverseEnd.Name));
            wiLinkTypes = wiLinkTypes.OrderBy(f => f).Distinct().ToList();
            wiLinkTypes.Insert(0, ResourceHelper.GetResourceString("ALL"));
            popup.Relateds.Items.Clear();
            popup.Relateds.ItemsSource = wiLinkTypes;
            popup.Relateds.SelectedIndex = 0;
        }

        /// <summary>
        /// Is Linked
        /// </summary>
        /// <param name="itemSource">
        /// </param>
        /// <param name="itemTarget">
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool IsLinked(WorkItem itemSource, WorkItem itemTarget)
        {
            if (itemSource.Id != itemTarget.Id && itemSource.Links.Count > 0)
            {
                foreach (WorkItemLink link in itemSource.WorkItemLinks)
                {
                    if (matrixReportData.RelatedSelectedItem != ResourceHelper.GetResourceString("ALL"))
                    {
                        if ((link.LinkTypeEnd.Name == matrixReportData.RelatedSelectedItem) && itemTarget.WorkItemLinks.Cast<WorkItemLink>().Any(l => l.SourceId == link.TargetId))
                        {
                            return true;
                        }
                    }
                    else if (itemTarget.WorkItemLinks.Cast<WorkItemLink>().Any(l => l.SourceId == link.TargetId))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Set Matrix Title
        /// </summary>
        /// <param name="projectName">
        /// </param>
        /// <param name="title">
        /// </param>
        private void SetMatrixTitle(string projectName, string title)
        {
            object styleTitle = WdBuiltinStyle.wdStyleTitle;
            Range parRange = OfficeHelper.CreateParagraphRange(ref document);
            parRange.Text = projectName + " " + title;
            parRange.set_Style(ref styleTitle);
            parRange.SetRange(parRange.Characters.First.Start, parRange.Characters.First.Start + projectName.Length);
            parRange.Select();
            switch (tfsManager.tfsVersion)
            {
                  case TfsVersion.Tfs2011:
                    {
                        parRange.Hyperlinks.Add(parRange, string.Format("{0}/{1}", tfsManager.collection.Name, projectName));
                    }
                    break;
                    case TfsVersion.Tfs2010:
                    {
                        parRange.Hyperlinks.Add(parRange, string.Format("{0}/web/Index.aspx?pguid={1}", tfsManager.GetTfsUrl(),tfsManager.GetProject(projectName).Guid));
                    }
                    break;
            }
            document.Content.InsertParagraphAfter();

            string DateFrom = string.Empty;
            string DateTo = string.Empty;

            if (matrixReportData.DateFrom != null)
                DateFrom = String.Format("{0}: {1}",WordToTFS.Properties.Resources.MatrixReport_Date_From,matrixReportData.DateFrom.Value.ToShortDateString());
            if (matrixReportData.DateTo != null)
                DateTo = String.Format("{0}: {1}",WordToTFS.Properties.Resources.MatrixReport_Date_To,matrixReportData.DateTo.Value.ToShortDateString());
            else
            {
                DateTo = String.Format("{0}: {1}",WordToTFS.Properties.Resources.MatrixReport_Date_To,DateTime.Now.ToShortDateString());
            }
            if (matrixReportData.DateFrom == null && matrixReportData.DateTo == null)
            {
                DateFrom = String.Format("{0}: {1}",WordToTFS.Properties.Resources.MatrixReport_Date,DateTime.Now.ToShortDateString());
                DateTo = string.Empty;
            }

            Range addInform = OfficeHelper.CreateParagraphRange(ref document);
            addInform.Font.Name = "Calibri";
            addInform.Font.Size = 11;
            addInform.Font.Color = WdColor.wdColorBlueGray;
            addInform.Font.Italic = -1;
            addInform.Text = String.Format("{0}: {1}             {2}          {3}",WordToTFS.Properties.Resources.MatrixReport_Link_Type,matrixReportData.RelatedSelectedItem,DateFrom,DateTo);
            addInform.Select();
            document.Content.InsertParagraphAfter();
    

        }

        #endregion
    }

    /// <summary>
    ///     The matrix data.
    /// </summary>
    public class MatrixData
    {
        #region Constructors and Destructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="MatrixData" /> class.
        /// </summary>
        public MatrixData()
        {
            Rows = new List<WorkItem>();
            Columns = new List<WorkItem>();
        }

        #endregion

        #region Public Properties

        /// <summary>
        ///     Gets or sets the columns.
        /// </summary>
        public List<WorkItem> Columns { get; set; }

        /// <summary>
        ///     Gets or sets the rows.
        /// </summary>
        public List<WorkItem> Rows { get; set; }

        #endregion
    }
}