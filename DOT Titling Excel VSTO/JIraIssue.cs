using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace DOT_Titling_Excel_VSTO
{
    class JiraIssue
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

        public async static Task<List<Issue>> GetAllStoriesAndBugs(string projectKey)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new System.Text.StringBuilder();
                jql.Append("project = " + projectKey);
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Software Bug\", Story)");
                jql.Append(" AND ");
                jql.Append("summary ~ \"!DELETE\"");
                List<Issue> filteredIssues = await FilterIssues(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }


        public async static Task<List<Issue>> GetAllTasks(string projectKey)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new System.Text.StringBuilder();
                jql.Append("project = " + projectKey);
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Task\")");
                jql.Append(" AND ");
                jql.Append("\"Epic Link\" = " + projectKey + "-945");
                List<Issue> filteredIssues = await FilterIssues(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<List<Issue>> GetAllEpics(string projectKey)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

                //Create the JQL
                var jql = new System.Text.StringBuilder();
                jql.Append("project = " + projectKey);
                jql.Append(" AND ");
                jql.Append("issuetype in (\"Epic\")");
                jql.Append(" AND ");
                jql.Append("summary ~ \"!DELETE\"");
                List<Issue> filteredIssues = await FilterIssues(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private static async Task<List<Issue>> FilterIssues(System.Text.StringBuilder jql)
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

        //public async static Task<List<Issue>> GetAllIssues(string projectKey, string ticketType)
        //{
        //    try
        //    {
        //        ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

        //        //Create the JQL
        //        var jql = new System.Text.StringBuilder();
        //        jql.Append("project = " + projectKey);

        //        if (ticketType == "Epics")
        //        {
        //            jql.Append(" AND ");
        //            jql.Append("issuetype in (\"Epic\")");
        //        }
        //        if (ticketType == "Tasks")
        //        {
        //            jql.Append(" AND ");
        //            jql.Append("issuetype in (\"Task\")");
        //            jql.Append(" AND ");
        //            jql.Append("\"Epic Link\" = " + projectKey + "-945");
        //        }
        //        if (ticketType == "Tickets")
        //        {
        //            jql.Append(" AND ");
        //            jql.Append("issuetype in (\"Software Bug\", Story)");
        //        }

        //        jql.Append(" AND ");
        //        jql.Append("summary ~ \"!DELETE\"");

        //        var issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize);
        //        var totalIssues = issues.TotalItems;
        //        var totalPages = (double)totalIssues / (double)ThisAddIn.PageSize;
        //        totalPages = Math.Ceiling(totalPages);
        //        var allIssues = issues.ToList();
        //        for (int currentPage = 1; currentPage < totalPages; currentPage++)
        //        {
        //            int startRecord = ThisAddIn.PageSize * currentPage;
        //            issues = await ThisAddIn.GlobalJira.Issues.GetIssuesFromJqlAsync(jql.ToString(), ThisAddIn.PageSize, startRecord);
        //            allIssues.AddRange(issues.ToList());
        //            if (issues.Count() == 0)
        //            {
        //                break;
        //            }
        //        }
        //        var filteredIssues = allIssues.Where(i =>
        //                    i.Summary != "DELETE").ToList();
        //        return filteredIssues;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error :" + ex);
        //        return null;
        //    }
        //}
        
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

        public static bool SaveSummary(string jiraId, string newValue, bool multiple)
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
                if (!multiple)
                    MessageBox.Show("Summary updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Summary could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveRelease(string jiraId, string newValue, bool multiple)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                string curRelease = ExtractRelease(issue);
                if (curRelease == newValue)
                {   if (!multiple)
                     MessageBox.Show("No change needed.");
                    return true;
                }

                // Remove all of the existing versions
                var oldVersions = issue.AffectsVersions.ToList();
                foreach (var oldVersion in oldVersions)
                {
                    issue.AffectsVersions.Remove(oldVersion);
                }

                if (newValue.Trim() != string.Empty)
                    issue.AffectsVersions.Add(newValue);

                issue.SaveChanges();
                if (!multiple)
                    MessageBox.Show("Release updated successfully updated.");
                return true;
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveLabels(string jiraId, string newValue, bool multiple)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                List<string> listofJiraLabels = ExtractListOfLabels(issue);
                List<string> listofExcelLabels = CreateListOfLabels(newValue);
                List<string> addLabels = listofExcelLabels.Except(listofJiraLabels).ToList();
                List<string> removeLabels = listofJiraLabels.Except(listofExcelLabels).ToList();

                if (addLabels.Count > 0)
                { 
                    foreach (string label in addLabels)
                    {
                        issue.Labels.Add(label);
                    }
                    issue.SaveChanges();
                }

                if (removeLabels.Count > 0)
                {
                    foreach (string label in removeLabels)
                    {
                        issue.Labels.Remove(label);
                    }
                    issue.SaveChanges();
                }


                return true;
                //string curRelease = ExtractRelease(issue);
                //if (curRelease == newValue)
                //{
                //    if (!multiple)
                //        MessageBox.Show("No change needed.");
                //    return true;
                //}

                //// Remove all of the existing versions
                //var oldVersions = issue.AffectsVersions.ToList();
                //foreach (var oldVersion in oldVersions)
                //{
                //    issue.AffectsVersions.Remove(oldVersion);
                //}

                //if (newValue.Trim() != string.Empty)
                //    issue.AffectsVersions.Add(newValue);

                //issue.SaveChanges();
                //if (!multiple)
                //    MessageBox.Show("Release updated successfully updated.");

            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveStatus(string jiraId, string newValue, bool multiple)
        {
            try
            {
                var issue = GetIssue(jiraId).Result;
                if (issue.Status.Name == newValue)
                {
                    if (!multiple)
                        MessageBox.Show("No change needed.");
                    return true;
                }
                issue.WorkflowTransitionAsync(newValue);
                if (!multiple)
                    MessageBox.Show("Status transitioned successfully.");
                return true;
            }
            catch
            {
                MessageBox.Show("Status could NOT be transitioned to " + newValue);
                return true;
            }
        }

        public static bool SaveCustomField(string jiraId, string field, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetIssue(jiraId).Result;
                if (issue[field] == newValue)
                {
                    if (!multiple)
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
                issue.SaveChanges(); if (!multiple)
                    MessageBox.Show(field + " successfully updated.");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                //MessageBox.Show(field + " could NOT successfully updated.");
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

        public static string ExtractLabels(Issue issue)
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

        public static List<string> CreateListOfLabels(string labels)
        {
            labels = labels.Replace(", ", ",");
            return labels.Split(',').ToList();
        }

        public static List<string> ExtractListOfLabels(Issue issue)
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
            item = item.Trim();
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
                case "issue.Assignee":
                    val = issue.Assignee;
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
                case "Labels":
                    List<string> listofLabels = ExtractListOfLabels(issue);
                    foreach (string label in listofLabels)
                    {
                        val = val + label + ", ";
                    }
                    if (val != string.Empty && val.Right(2) == ", ")
                        val = val.Left(val.Length - 2);                        
                    break;
                default:
                    break;
            }
            return val;
        }
    }
}
