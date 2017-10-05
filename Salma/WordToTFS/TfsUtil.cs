using System;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Build.Client;

namespace WordToTFS
{ 
    /// <summary>
    /// Responsible for indicating TFS version
    /// </summary>
    public static class TfsUtil
    {
        /// <summary>
        /// Indicates wich TFS version we work with
        /// </summary>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static TfsVersion GetVersion(this TfsTeamProjectCollection collection)
        {
            if (collection == null)
                throw new ArgumentNullException("collection");
            try
            {

                IBuildServer buildserver = collection.GetService<IBuildServer>();
                string ServerVersion = Convert.ToString(buildserver.BuildServerVersion);

                switch (ServerVersion)
                {

                    case "V5":
                        {
                            return TfsVersion.Tfs2011;
                            break;
                        }
                    case "V4":
                        {
                            return TfsVersion.Tfs2011;
                            break;
                        }
                    case "V3":
                        {
                            return TfsVersion.Tfs2010;
                            break;
                        }
                }
                ITestManagementService testService = collection.GetService<ITestManagementService>();
                TfsTeamService teamService = collection.GetService<TfsTeamService>();
                teamService.QueryTeams(string.Empty);
                return TfsVersion.Tfs2011;
            }
            catch
            {
                return TfsVersion.Tfs2010;
            }
        }
    }
    /// <summary>
    /// Tfs versions.
    /// </summary>
   public enum TfsVersion
    {
        Tfs2008,
        Tfs2010,
        Tfs2011,
        Tfs2013
    }
}
 
    

