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
        public async static Task<Issue> GetIssue(string jiraId)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var issue = await ThisAddIn.GlobalJira.Issues.GetIssueAsync(jiraId);
                return issue;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        [Obsolete("GetIssueWithLinq is deprecated, please use GetIssue instead.")]
        public static List<Issue> GetIssueWithLinq(string jiraId)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
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

        public async static Task<List<Issue>> GetAllIssues()
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new System.Text.StringBuilder();
                jql.Append("project = DOTTITLNG");
                jql.Append(" && ");
                jql.Append("issuetype in (\"Software Bug\", Story)");
                jql.Append(" && ");
                jql.Append("summary ~ \"!DELETE\"");

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
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        [Obsolete("GetAllIssuesWithLinq is deprecated, please use GetAllIssues instead.")]
        public static List<Issue> GetAllIssuesWithLinq()
        {
            ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
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

        public async static Task<IDictionary<string, Issue>> GetSelectedIssues(params string[] jiraIDs)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesAsync(jiraIDs);
                return issues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }
        public static void SaveSummary(string jiraId, string newSummary)
        {
            var issue = GetIssue(jiraId).Result;
            issue.Summary = newSummary;
            issue.SaveChanges();
        }
    }
}
