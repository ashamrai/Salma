
namespace WordToTFS
{
    public class SalmaConstants
    {
        /// <summary>
        /// Constants for Word Document Comments
        /// </summary>
        public static class Comments
        {
            public static readonly string WI_ID = ResourceHelper.GetResourceString("WI_ID");
            public static readonly string FOR_WI_ID = ResourceHelper.GetResourceString("FOR_WI_ID");
            public static readonly string CREATED_BY = "Created by:";
            public static readonly string TYPE = "Type:";
            public static readonly string TITLE = "Title:";
            public static readonly string STATUS = "Status:";
            public static readonly string PROJECT = "Project:";
            public static readonly string FOR = ResourceHelper.GetResourceString("FOR");

            
        }
        public static class TFS
        {
           
            public const string PROJECT_SERVER_SYNC_ASSIGMENT_DATA = "Project Server Sync Assignment Data";
            public const string LOCAL_DATA_SOURCE = "Local Data Source";
        }       
    }
}
