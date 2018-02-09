using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Text;
using Jira = Atlassian.Jira;

//// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
namespace DOT_Titling_Excel_VSTO
{
    class JiraShared
    {
        //Get From Jira
        public async static Task<Jira.Issue> GetSingleFromJira(string issueID)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var issue = await ThisAddIn.GlobalJira.Issues.GetIssueAsync(issueID);
                return issue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<List<Jira.Issue>> GetSelectedFromJira(List<string> listofIssueIDs)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("key in (");
                jql.Append(FormatListofIDs(listofIssueIDs));
                jql.Append(")");
                List<Jira.Issue> filteredIssues = await JiraShared.Filter(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private async static Task<IDictionary<string, Jira.Issue>> GetSelectedFromJiraAlternative(params string[] listofIssueIDs)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesAsync(listofIssueIDs);
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static async Task<List<Jira.Issue>> Filter(StringBuilder jql)
        {
            var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize);
            var totalIssues = issues.TotalItems;
            var totalPages = (double)totalIssues / (double)ThisAddIn.PageSize;
            totalPages = Math.Ceiling(totalPages);
            var allIssues = issues.ToList();
            for (int currentPage = 1; currentPage < totalPages; currentPage++)
            {
                int startRecord = ThisAddIn.PageSize * currentPage;
                issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize, startRecord);
                allIssues.AddRange(issues.ToList());
                if (issues.Count() == 0)
                {
                    break;
                }
            }
            var filteredIssues = allIssues.Where(i =>
                        i.Summary != "DELETE").ToList();
            return filteredIssues;
        }

        public static List<string> CreateListOfLabels(string labels)
        {
            labels = labels.Replace(", ", ",");
            return labels.Split(',').ToList();
        }

        public static StringBuilder FormatProjectList(List<string> listofProjects)
        {
            var projectList = new StringBuilder();
            int cnt = 1;
            int projectCount = listofProjects.Count();
            foreach (string project in listofProjects)
            {
                projectList.Append(project);
                if (cnt != projectCount)
                    projectList.Append(", ");

                cnt++;
            }
            return projectList;
        }

        public static StringBuilder FormatListofIDs(List<string> lst)
        {
            var idList = new StringBuilder();
            int cnt = 1;
            int iCnt = lst.Count();
            foreach (string project in lst)
            {
                idList.Append(project);
                if (cnt != iCnt)
                    idList.Append(", ");
                cnt++;
            }
            return idList;
        }

        //Extract
        public static string ExtractRelease(Jira.Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.AffectsVersions)
            {
                if (c > 0)
                    val = val + "; ";
                val = val + issue.AffectsVersions[c].Name;
                c++;
            }
            return val;
        }

        public static string ExtractLabels(Jira.Issue issue)
        {
            string val = string.Empty;
            if (issue.Labels.Count > 0)
            {
                foreach (var label in issue.Labels)
                {
                    val = val + "[" + label + "]";
                }
            }
            return val;
        }

        public static List<string> ExtractListOfLabels(Jira.Issue issue)
        {
            List<string> listofLabels = new List<string>();
            if (issue.Labels.Count > 0)
            {
                foreach (var label in issue.Labels)
                {
                    listofLabels.Add(label);
                }
            }
            return listofLabels;
        }

    }
}
