﻿using System;
using Jira = Atlassian.Jira;

namespace DOT_Titling_Excel_VSTO
{
    public partial class ThisAddIn
    {
        public static string DesktopDir = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public static string OutputDir = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Out";
        public static string InputDir = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\In";
        public static string RoadMapDir = @Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Road Map";
        public static string JiraSite = "https://wiportal.atlassian.net";
        public static int MaxJiraRequests = 1000;
        public static int MaxRecordsToProcess = 100000;
        public static int PageSize = 100;
        public static string R3Folder = "https://wisdot.sharepoint.com/sites/bitsproj/3025/SitePages/Home.aspx?RootFolder=%2Fsites%2Fbitsproj%2F3025%2FProject%20Documents%2FDMV-BVS-DAS%20Project%20Team%20Docs%2FRelease%203&FolderCTID=0x0120008B2FF5906472224CB44A99C8A95ADAF9&View=%7BBB5263F9-5D41-4EC4-9518-EAF825B2CB19%7D";
        public static string ProjectKeyDOT = "DOTTITLNG";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
