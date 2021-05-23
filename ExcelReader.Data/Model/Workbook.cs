using System.Collections.Generic;

namespace ExcelReader.Data.Model
{
    public class Workbook
    {
        public string Name { get; set; }
        public List<Worksheet> Worksheets { get; set; }

        public Workbook(string name)
        {
            this.Name = name;
            this.Worksheets = new List<Worksheet>();
        }
    }
}
