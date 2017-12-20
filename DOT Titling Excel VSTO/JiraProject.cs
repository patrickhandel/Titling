using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace DOT_Titling_Excel_VSTO
{
    class JiraProject
    {
        public async static Task<Project> GetProject(string jiraId)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var project = await ThisAddIn.GlobalJira.Projects.GetProjectAsync(jiraId);
                return project;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        //public async static Task<List<Project>> GetAllProjects()
        //{
        //    try
        //    {
        //        ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;

        //        //Create the JQL
        //        var jql = new System.Text.StringBuilder();
        //        jql.Append("project = DOTTITLNG");
        //        jql.Append(" && ");
        //        if (type == "Tickets")
        //        {
        //            jql.Append("issuetype in (\"Software Bug\", Story)");
        //        }

        //        jql.Append(" && ");
        //        jql.Append("summary ~ \"!DELETE\"");

        //        var projects = await ThisAddIn.GlobalJira.Projects.GetProjectsAsync(jql.ToString(), ThisAddIn.PageSize);
        //        var totalProjects = projects.TotalItems;
        //        var totalPages = (double)totalProjects / (double)ThisAddIn.PageSize;
        //        totalPages = Math.Ceiling(totalPages);
        //        var allIssues = projects.ToList();
        //        for (int currentPage = 1; currentPage < totalPages; currentPage++)
        //        {
        //            int startRecord = ThisAddIn.PageSize * currentPage;
        //            projects = await ThisAddIn.GlobalJira.Projects.GetProjectsFromJqlAsync(jql.ToString(), ThisAddIn.PageSize, startRecord);
        //            allIssues.AddRange(projects.ToList());
        //            if (projects.Count() == 0)
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
        
        //public async static Task<IDictionary<string, Project>> GetSelectedProjects(params string[] jiraIDs)
        //{
        //    try
        //    {
        //        ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = ThisAddIn.MaxJiraRequests;
        //        var projects = await ThisAddIn.GlobalJira.Projects.GetProjectsAsync(jiraIDs);
        //        return projects;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error :" + ex);
        //        return null;
        //    }
        //}

        //public static bool SaveName(string jiraId, string newValue)
        //{
        //    try
        //    {
        //        var project = GetProject(jiraId).Result;
        //        if (project.Name == newValue)
        //        {
        //            MessageBox.Show("No change needed.");
        //            return true;
        //        }
        //        project.Name = newValue;
        //        project.SaveChanges();
        //        MessageBox.Show("Summary updated successfully updated.");
        //        return true;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Summary could NOT successfully updated.");
        //        return false;
        //    }
        //}

    }
}
