using ExcelDataReader;
using ExcelReader.Data.Model;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelReader.Data.Loader
{
    public class ExcelLoaderV2
    {
        public List<Database> LoadAll(List<string> filePaths)
        {
            List<Database> server = new List<Database>();

            foreach(string filePath in filePaths)
            {
                server.Add(Load(filePath));
            }

            return server;
        }

        public Database Load(string filePath)
        {
            Database database = null;

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    database = new Database(filePath);

                    for (int i = 0; i < result.Tables.Count; i++)
                    {
                        DataTable dataTable = result.Tables[i];

                        var columns = new List<Column>();
                        for (int ci = 0; ci < dataTable.Columns.Count; ci++)
                        {
                            DataColumn dataColumn = dataTable.Columns[ci];

                            if (dataColumn.ColumnName.StartsWith("Column"))
                                break;

                            Column column = new Column(ci, dataColumn.ColumnName);
                            columns.Add(column);
                        }

                        Table table = new Table(i, dataTable.TableName, columns);

                        for (int ri = 0; ri < dataTable.Rows.Count; ri++)
                        {
                            DataRow dataRow = dataTable.Rows[ri];

                            Row row = new Row(ri);

                            for (int ci = 0; ci < columns.Count; ci++)
                            {
                                Cell cell = new Cell(columns[ci], dataRow.ItemArray[ci]);
                                row.Cells.Add(cell);
                            }

                            table.Rows.Add(row);
                        }

                        database.Tables.Add(table);
                    }
                }
            }

            return database;
        }
    }
}
