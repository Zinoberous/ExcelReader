using System.Collections.Generic;

namespace ExcelReader.Data.Model
{
    public class Worksheet
    {
        public string Name { get; set; }
        public List<Column> Columns { get; set; }
        public List<Row> Rows { get; set; }

        public Worksheet(string name, List<Column> columns)
        {
            this.Name = name;
            this.Columns = columns;
            this.Rows = new List<Row>();
        }
    }
}
