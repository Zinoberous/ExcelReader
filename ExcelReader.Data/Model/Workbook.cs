using System.Collections.Generic;

namespace ExcelReader.Data.Model
{
    public class Workbook
    {
        public string FilePath { get; }
        public string Name { get; set; }
        public List<Worksheet> Worksheets { get; set; }

        public Workbook(string filePath)
        {
            this.FilePath = filePath;
            this.Name = filePath.Split('\\')[filePath.Split('\\').Length - 1].Split('.')[0];
            this.Worksheets = new List<Worksheet>();
        }
    }
}
