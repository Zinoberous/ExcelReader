using System.Collections.Generic;

namespace ExcelReader.Data.Model
{
    public class Row
    {
        public int Index { get; set; }
        public List<Cell> Cells { get; set; }

        public Row(int index)
        {
            this.Index = index;
            this.Cells = new List<Cell>();
        }
    }
}
