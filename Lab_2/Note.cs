using System.Collections.Generic;

namespace Lab_2
{
    class Note
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
        public Note(List<string> args)
        {
            Properties = new Dictionary<string, string>();
            int j = 0;
            foreach (string item in PropertiesSet)
            {
                Properties[item] = args[j];
                j++;
            }
        }

        public Note()
        {
            
        }
        /*public override string ToString()
        {
            string result = "";
            foreach (var item in Properties)
            {
                result += $"{item.Key}: {item.Value}" + "\n";
            }
            return result;
        }*/
    }
}