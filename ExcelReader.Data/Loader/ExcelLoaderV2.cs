using ExcelDataReader;
using ExcelReader.Data.Model;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelReader.Data.Loader
{
    public class ExcelLoaderV2
    {
        public List<Workbook> LoadAll(List<string> filePaths)
        {
            List<Workbook> workbooks = new List<Workbook>();

            foreach(string filePath in filePaths)
            {
                workbooks.Add(Load(filePath));
            }

            return workbooks;
        }

        public Workbook Load(string filePath)
        {
            Workbook workbook = null;

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

                    workbook = new Workbook(filePath);

                    for (int i = 0; i < result.Tables.Count; i++)
                    {
                        DataTable dataTable = result.Tables[i];

                        var columns = new List<Column>();
                        for (int ci = 0; ci < dataTable.Columns.Count; ci++)
                        {
                            DataColumn dataColumn = dataTable.Columns[ci];
                            columns.Add(new Column(ci, dataColumn.ColumnName));
                        }

                        var worksheet = new Worksheet(i, dataTable.TableName, columns);

                        for (int ri = 0; ri < dataTable.Rows.Count; ri++)
                        {
                            DataRow dataRow = dataTable.Rows[ri];

                            Row row = new Row(ri)
                            {
                                Cells = new List<Cell>()
                            };

                            for (int ci = 0; ci < dataRow.ItemArray.Length; ci++)
                            {
                                row.Cells.Add(new Cell(columns[ci], dataRow.ItemArray[ci]));
                            }

                            worksheet.Rows.Add(row);
                        }

                        workbook.Worksheets.Add(worksheet);
                    }
                }
            }

            return workbook;
        }
    }
}
