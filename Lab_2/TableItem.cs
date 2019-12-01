using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_2
{
    class TableItem
    {
        public static HashSet<string> PropertiesSet = new HashSet<string>()
        {
            "ID",
            "Наименование",
            "Описание",
            "Источник",
            "Объект воздействия",
            "Нарушение конфиденциальности",
            "Нарушение целостности",
            "Нарушение доступности"
        };

        public Dictionary<string, string> Properties { get; set; }
        /*
                public string Id { set; get; }
                public string Name { set; get; }
                public string Description { set; get; }
                public string Source { set; get; }
                public string ThreatObject { set; get; }
                public bool PrivacyBreach { set; get; }
                public bool IntegrityBreach { set; get; }
                public bool AccessBreach { set; get; }*/


        public TableItem(List<string> args)
        {
            Properties = new Dictionary<string, string>();
            int j = 0;
            foreach (string item in PropertiesSet)
            {
                Properties[item] = args[j];
                j++; 
            }
            /*Id = args[0];
            Name = ReplaceBreakLines(args[1]);
            Description = *//*ReplaceBreakLines(args[2]);
            Source = args[3];
            ThreatObject = args[4];
            PrivacyBreach = (args[5] == "1");
            IntegrityBreach = (args[6] == "1");
            AccessBreach = (args[7] == "1");*/
        }

        public override string ToString()
        {
            string result = "";
            foreach (var item in Properties)
            {
                result += $"{item.Key}: {item.Value}";
            }
            return ReplaceBreakLines(result);
        }

        public static string ReplaceBreakLines(string str)
        {
            return str.Replace("_x000d_", "").Replace("\n", "");
        }
    }

}