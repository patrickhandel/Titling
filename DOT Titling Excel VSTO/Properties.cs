using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Newtonsoft.Json;

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

    public static class WorksheetPropertiesManager
    {
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
    }
}
