using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DOT_Titling_Excel_VSTO
{
    class JiraUtils
    {
        public static List<Issue> GetSingleIssue(string jiraId)
        {
            try
            {
                ThisAddIn.GlobalJira.MaxIssuesPerRequest = 1;
                var issues = (from i in ThisAddIn.GlobalJira.Issues.Queryable
                              where i.Key == jiraId
                              select i).ToList();
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static List<Issue> GetAllIssues()
        {
            ThisAddIn.GlobalJira.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
            var issues = (from i in ThisAddIn.GlobalJira.Issues.Queryable
                          where i.Project == "DOTTITLNG" &&
                          (i.Type == "Story" || i.Type == "Software Bug") &&
                          i.Summary != "DELETE"
                          orderby i.Created
                          select i).ToList();

            // try this: i.Summary == new LiteralMatch("My Title")

            var issuesToRemove = issues.FindAll(x => x.Summary.ToUpper().Trim() == "DELETE");
            foreach (var issueToRemove in issuesToRemove)
            {
                issues.Remove(issues.FirstOrDefault(x => x.Key.Value == issueToRemove.Key.Value));
            }
            return issues;
        }

        public async static Task<Issue> GetIssue(string jiraId)
        {
            try
            {
                ThisAddIn.GlobalJira.MaxIssuesPerRequest = 1;
                var issue = await ThisAddIn.GlobalJira.Issues.GetIssueAsync(jiraId);
                return issue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<IDictionary<string, Issue>> GetIssues(params string[] IDs)
        {
            try
            {
                ThisAddIn.GlobalJira.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesAsync(IDs);
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<IPagedQueryResult<Issue>> GetTitlingIssues()
        {
            try
            {
                ThisAddIn.GlobalJira.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                var jql = new System.Text.StringBuilder();
                jql.Append("project = DOTTITLNG");
                var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString());
                var filteredIssues = issues.Where(i =>
                            i.Summary != "DELETE" &&
                            (i.Type == "Story" || i.Type == "Software Bug"));
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }
    }
}
