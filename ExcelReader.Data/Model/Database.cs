using System.Collections.Generic;

namespace ExcelReader.Data.Model
{
    public class Database
    {
        public string FilePath { get; }
        public string Name { get; set; }
        public List<Table> Tables { get; set; }

        public Database(string filePath)
        {
            this.FilePath = filePath;
            this.Name = filePath.Split('\\')[filePath.Split('\\').Length - 1].Split('.')[0];
            this.Tables = new List<Table>();
        }
    }
}
