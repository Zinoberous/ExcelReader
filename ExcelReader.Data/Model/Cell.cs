namespace ExcelReader.Data.Model
{
    public class Cell
    {
        public Column Column { get; set; }
        public object Value { get; set; }

        public Cell(Column column, object value)
        {
            this.Column = column;
            this.Value = value;
        }
    }
}
