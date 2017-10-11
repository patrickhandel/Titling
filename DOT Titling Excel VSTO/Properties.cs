using System;
using System.Collections.Generic;
using System.Configuration;
using Newtonsoft.Json;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DOT_Titling_Excel_VSTO
{
    public class WorksheetProperties
    {
        public string Worksheet { get; set; }

        public string Range { get; set; }
    }

    public class ColumnTypes
    {
        public string Name { get; set; }

        public int Width { get; set; }
    }

    public class MailMergeFields
    {
        public string Name { get; set; }

        public string Text { get; set; }
    }

    public class JiraFields
    {
        public string Range { get; set; }

        public string ColumnHeader { get; set;  }

        public string Type { get; set; }

        public string Value { get; set; }

        public string Formula { get; set; }
    }

    public static class WorksheetPropertiesManager
    {
        public static List<JiraFields> GetJiraFields(Excel.Worksheet ws)
        {
            var str = ConfigurationManager.AppSettings["JiraFields"];
            List<JiraFields> lst = JsonConvert.DeserializeObject<List<JiraFields>>(str);
            string range = (ws.Name == "Epics") ? "EpicData" : "TicketData";
            return lst.FindAll(y => y.Range == range);
        }

        public static  List<WorksheetProperties> GetWorksheetProperties()
        {
            var str = ConfigurationManager.AppSettings["WorksheetProperties"];
            List<WorksheetProperties> lst = JsonConvert.DeserializeObject<List<WorksheetProperties>>(str);
            return lst;
        }

        public static List<ColumnTypes> GetColumnTypes()
        {
            var str = ConfigurationManager.AppSettings["ColumnTypes"];
            List<ColumnTypes> lst = JsonConvert.DeserializeObject<List<ColumnTypes>>(str);
            return lst;
        }

        public static List<MailMergeFields> GetMailMergeFields()
        {
            try
            {
                var str = ConfigurationManager.AppSettings["MailMergeFields"];
                List<MailMergeFields> lst = JsonConvert.DeserializeObject<List<MailMergeFields>>(str);
                return lst;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return null;
            }
        }
    }
}
