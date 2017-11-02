using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.Linq;
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

        //public async static Task<List<IssueChangeLog>> GetChangeLog(string jiraId)
        //{
        //    try
        //    {
        //        ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1000;
        //        IssueChangeLog changes = await ThisAddIn.GlobalJira.Issues.GetChangeLogsAsync(jiraId);
        //        return changes;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error :" + ex);
        //        return null;
        //    }
        //}


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

        public async static Task<List<Issue>> GetAllIssues(string type = "Tickets")
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new System.Text.StringBuilder();
                jql.Append("project = DOTTITLNG");
                jql.Append(" && ");

                if (type == "Epics")
                {
                    jql.Append("issuetype in (\"Epic\")");
                }
                if (type == "Tickets")
                {
                    jql.Append("issuetype in (\"Software Bug\", Story)");
                }

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

        public static bool SaveSummary(string jiraId, string newValue)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                if (issue.Summary == newValue)
                {
                    MessageBox.Show("No change needed.");
                    return true;
                }
                issue.Summary = newValue;
                issue.SaveChanges();
                MessageBox.Show("Summary updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Summary could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveRelease(string jiraId, string newValue)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                string curRelease = ExtractRelease(issue);
                if (curRelease == newValue)
                {
                    MessageBox.Show("No change needed.");
                    return true;
                }

                // Remove all of the existing versions
                var oldVersions = issue.AffectsVersions.ToList();
                foreach (var oldVersion in oldVersions)
                {
                    issue.AffectsVersions.Remove(oldVersion);
                }

                // Add thew new version
                issue.AffectsVersions.Add(newValue);

                issue.SaveChanges();
                MessageBox.Show("Release updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveStatus(string jiraId, string newValue)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                if (issue.Status.Name == newValue)
                {
                    MessageBox.Show("No change needed.");
                    return true;
                }
                issue.WorkflowTransitionAsync(newValue);
                MessageBox.Show("Status transitioned successfully.");
                return true;
            }
            catch
            {
                MessageBox.Show("Status could NOT be transitioned to " + newValue);
                return true;
            }
        }

        public static bool SaveCustomField(string jiraId, string field, string newValue)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetIssue(jiraId).Result;
                if (issue[field] == newValue)
                {
                    MessageBox.Show("No change needed.");
                    return true;
                }
                if (newValue == string.Empty)
                {
                    issue[field] = null;
                }
                else
                {
                    issue[field] = newValue;
                }
                issue.SaveChanges();
                MessageBox.Show(field + " successfully updated.");
                return true;
            }
            catch //(Exception ex)
            {
                //MessageBox.Show("Error :" + ex);
                MessageBox.Show(field + " could NOT successfully updated.");
                return false;
            }
        }

        public static string ExtractRelease(Issue issue)
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

        public static string ExtractFixRelease(Issue issue)
        {
            string val = string.Empty;
            int c = 0;
            foreach (var ver in issue.FixVersions)
            {
                val = issue.FixVersions[c].Name;
                c++;
            }
            return val;
        }

        public static string ExtractDOTWebServices(Issue issue)
        {
            string val = string.Empty;
            if (issue["DOT Web Services"] != null)
            {
                foreach (var ver in issue.CustomFields["DOT Web Services"].Values)
                {
                    val = val + " " + ver;
                }
                val = val.Trim().Replace(" ", ", ");
            }
            return val;
        }

        public static string ExtractSprintNumber(Issue issue)
        {
            //string val = ExtractCustomValue(issue, "Sprint");
            string val = string.Empty;
            int thisSprint = 0;
            int lastSprint = 0;
            foreach (var value in issue.CustomFields["Sprint"].Values)
            {
                val = value;
                //val = val.Replace("DOT", "");
                //val = val.Replace("Backlog", "");
                //val = val.Replace("Hufflepuff", "");
                //val = val.Replace("for", "");
                //val = val.Replace("Sprint", "");
                //val = val.Replace("Ready", "");
                //val = val.Replace("Other", "");
                //val = val.Replace("Approved", "");
                //val = val.Replace("-", "");
                //val = val.Replace(" ", "");
                //for (int rev = 1; rev <= 20; rev++)
                //    val = val.Replace("R" + rev.ToString(), "");

                if (val.Length > 2)
                {
                    val = val.Substring(val.Length - 3).Trim();
                    if (val != string.Empty)
                    {
                        if (Int32.TryParse(val, out thisSprint))
                        {
                            if (thisSprint > lastSprint)
                                lastSprint = thisSprint;
                        }
                    }
                }
            }

            string retval = string.Empty;
            if (lastSprint == 0)
            {
                retval = "";
            }
            else
            {
                retval = lastSprint.ToString();
            }
            return retval;
        }

        public static string ExtractCustomValue(Issue issue, string item)
        {
            string val = string.Empty;
            item = item.Replace(" Id ", " I'd ");
            try
            {
                val = issue[item].Value;
            }
            catch
            {
                val = string.Empty;
            }
            return val;
        }

        public static string ExtractStandardValue(Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "issue.Type.Name":
                    val = issue.Type.Name;
                    break;
                case "issue.Key.Value":
                    val = issue.Key.Value;
                    break;
                case "issue.Summary":
                    val = issue.Summary;
                    break;
                case "issue.Status.Name":
                    val = issue.Status.Name;
                    break;
                case "issue.Description":
                    val = issue.Description;
                    break;
                default:
                    break;
            }
            return val;
        }

        public static string ExtractValueBasedOnFunction(Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "Sprint":
                    val = ExtractSprintNumber(issue);
                    break;
                case "Release":
                    val = ExtractRelease(issue);
                    break;
                case "Fix Release":
                    val = ExtractFixRelease(issue);
                    break;
                case "DOT Web Services":
                    val = ExtractDOTWebServices(issue);
                    break;
                default:
                    break;
            }
            return val;
        }
    }
}
