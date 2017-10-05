using System;
using System.Collections.Generic;

namespace WordToTFS.Model
{
    /// <summary>
    /// Represent MatrixReport dialog data.
    /// </summary>
   internal class MatrixReportData
    {
        /// <summary>
        /// Columns types selected item.
        /// </summary>
        internal string ColumnsTypesSelectedItem { get; set; }

        /// <summary>
        /// Date from.
        /// </summary>
        internal DateTime? DateFrom { get; set; }

        /// <summary>
        /// Date to.
        /// </summary>
        internal DateTime? DateTo { get; set; }

        /// <summary>
        /// Include not linked items.
        /// </summary>
        internal bool IncludeNotLinkedItems { get; set; }

        /// <summary>
        /// Related selected item.
        /// </summary>
        internal string RelatedSelectedItem { get; set; }

        /// <summary>
        /// Rows types selected item.
        /// </summary>
        internal string RowsTypesSelectedItem { get; set; }

        /// <summary>
        /// State columns selected item.
        /// </summary>
        internal List<string> StateColumnsSelectedItem { get; set; }

        /// <summary>
        /// State rows selected item.
        /// </summary>
        internal List<string> StateRowsSelectedItem { get; set; }
    }
}
