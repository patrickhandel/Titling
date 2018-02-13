using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Text;
using Jira = Atlassian.Jira;
using Excel = Microsoft.Office.Interop.Excel;

//// https://bitbucket.org/farmas/atlassian.net-sdk/wiki/Home
namespace DOT_Titling_Excel_VSTO
{
    class JiraShared
    {
        // Enums
        public enum ImportType
        {
            AllIssues = 1,
            StoriesAndBugsOnly = 2,
            EpicsOnly = 3,
            TasksOnly = 4,
            ChecklistTasksOnly = 5
        };

        //Get From Jira
        public async static Task<List<Jira.Issue>> GetAllFromJira(List<string> listofProjects, ImportType importType)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
                //Create the JQL
                var jql = new StringBuilder();
                jql.Append("project in (");
                jql.Append(FormatProjectList(listofProjects));
                jql.Append(")");
                string jqlIssueTypes = GetJQLForImportType(importType);
                jql.Append(jqlIssueTypes);

                ////Handle DOT specifically
                //if (listofProjects[0] == ThisAddIn.ProjectKeyDOT)
                //{
                //    jql.Append(" AND ");
                //    jql.Append("issuetype in (\"Software Bug\", Story)");
                //}

                jql.Append(" AND ");
                jql.Append("summary ~ \"!DELETE\"");

                List<Jira.Issue> filteredIssues = await Filter(jql);
                return filteredIssues;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        private static string GetJQLForImportType(ImportType importType)
        {
            string jql = string.Empty;
            switch (importType)
            {
                case ImportType.AllIssues:
                    break;
                case ImportType.StoriesAndBugsOnly:
                    jql = " AND issuetype in (\"Software Bug\", Story)";
                    break;
                case ImportType.EpicsOnly:
                    jql = " AND issuetype in (\"Epic\")";
                    break;
                case ImportType.TasksOnly:
                    jql = " AND issuetype in (\"Task\")";
                    break;
                case ImportType.ChecklistTasksOnly:
                    //TO DO
                    break;
                default:
                    break;
            }
            return jql;
        }

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
                List<Jira.Issue> filteredIssues = await Filter(jql);
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

        public static bool SaveCustomField(string issueID, string field, string newValue, bool multiple)
        {
            try
            {
                newValue = newValue.Trim();
                var issue = GetSingleFromJira(issueID).Result;
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

        //Save Single Value
        public static bool SaveSummary(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
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
                MessageBox.Show("Summary could NOT be successfully updated.");
                return false;
            }
        }

        public static bool SaveRelease(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
                string curRelease = ExtractRelease(issue);
                if (curRelease == newValue)
                {
                    if (!multiple)
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

        public static bool SaveLabels(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
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
            }
            catch
            {
                MessageBox.Show("Release could NOT successfully updated.");
                return false;
            }
        }

        public static bool SaveStatus(string issueID, string newValue, bool multiple)
        {
            try
            {
                var issue = GetSingleFromJira(issueID).Result;
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

        //Extraction Methods
        public static string ExtractStandardValue(Jira.Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "issue.Project":
                    val = issue.Project;
                    break;
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

        public static string ExtractCustomValue(Jira.Issue issue, string item)
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

        public static string ExtractValueBasedOnFunction(Jira.Issue issue, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "Sprint Number":
                    val = ExtractSprintNumber(issue);
                    break;
                case "Release":
                    val = ExtractRelease(issue);
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

        //Extraction Functions
        private static string ExtractDOTWebServices(Jira.Issue issue)
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

        private static string ExtractSprintNumber(Jira.Issue issue)
        {
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

            string sprintNumber = string.Empty;
            if (lastSprint == 0)
            {
                sprintNumber = "";
            }
            else
            {
                sprintNumber = lastSprint.ToString();
            }
            return sprintNumber;
        }

        //Lists
        public static List<string> CreateListOfLabels(string labels)
        {
            labels = labels.Replace(", ", ",");
            return labels.Split(',').ToList();
        }

        public static List<Jira.Issue> GetListofSelectedIssuesIDsFromTable(Excel.Worksheet ws, Excel.Range selection)
        {
            List<string> listofIssues = new List<string>();
            for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
            {
                if (ws.Rows[row].EntireRow.Height != 0)
                {
                    int issueIDCol = SSUtils.GetColumnFromHeader(ws, "Issue ID");
                    string issueID = SSUtils.GetCellValue(ws, row, issueIDCol).Trim();
                    listofIssues.Add(issueID);
                }
            }
            var issues = GetSelectedFromJira(listofIssues).Result;
            return issues;
        }

        //String Builders
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


    }
}
