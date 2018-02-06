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
        public async static Task<Project> GetProject(string issueID)
        {
            try
            {
                ThisAddIn.GlobalJira.Issues.MaxIssuesPerRequest = 1;
                var project = await ThisAddIn.GlobalJira.Projects.GetProjectAsync(issueID);
                return project;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public async static Task<List<Project>> GetAllProjects()
        {
            try
            {
                var projects = await ThisAddIn.GlobalJira.Projects.GetProjectsAsync();
                return projects.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }

        public static string ExtractStandardValue(Project project, string item)
        {
            string val = string.Empty;
            switch (item)
            {
                case "Project.Name":
                    val = project.Name;
                    break;
                case "Project.Key":
                    val = project.Key;
                    break;
                case "Project.Id":
                    val = project.Id;
                    break;
                case "Project.Lead":
                    val = project.Lead;
                    break;
                default:
                    break;
            }
            return val;
        }

    }
}
