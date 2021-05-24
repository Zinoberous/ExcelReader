namespace ExcelReader.Data.Model
{
    public class Column
    {
        public int Index { get; set; }
        public string Name { get; set; }

        public Column(int index, string name)
        {
            this.Index = index;
            this.Name = name;
        }
    }
}
