using System.Collections.Generic;

namespace Lab_2
{
    class Threat
    {
        private static HashSet<string> PropertiesSet = new HashSet<string>()
        {
            "ID",
            "Наименование",
            "Описание",
            "Источник",
            "Объект",
            "Нарушение приватности",
            "Нарушение целостности",
            "Нарушение доступа"
        };

        public Dictionary<string, string> Properties { get; set; }
        public Threat(List<string> args)
        {
            Properties = new Dictionary<string, string>();
            int j = 0;
            foreach (string item in PropertiesSet)
            {
                Properties[item] = args[j];
                j++;
            }
        }
    }
}