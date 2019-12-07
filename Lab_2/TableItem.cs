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

        public TableItem(List<string> args)
        {
            Properties = new Dictionary<string, string>();
            int j = 0;
            foreach (string item in PropertiesSet)
            {
                Properties[item] = args[j];
                j++;
            }
        }

        public static string ReplaceBreakLines(string str)
        {
            return str.Replace("_x000d_", "").Replace("\n", "");
        }
    }

}