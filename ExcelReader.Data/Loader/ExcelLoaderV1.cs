using ExcelDataReader;
using ExcelReader.Data.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelReader.Data.Loader
{
    public class ExcelLoaderV1
    {
        public List<Dictionary<string, List<List<Tuple<string, object>>>>> LoadAll(List<string> filePaths)
        {
            var workbooks = new List<Dictionary<string, List<List<Tuple<string, object>>>>>();

            foreach(string filePath in filePaths)
            {
                workbooks.Add(Load(filePath));
            }

            return workbooks;
        }

        // Workbook<TableName, Rows<ColumnName, Values>>
        public Dictionary<string, List<List<Tuple<string, object>>>> Load(string filePath)
        {
            var workbook = new Dictionary<string, List<List<Tuple<string, object>>>>();

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

                    foreach (DataTable table in result.Tables)
                    {
                        var worksheet = new KeyValuePair<string, List<List<Tuple<string, object>>>>(table.TableName, new List<List<Tuple<string, object>>>());

                        var columnNames = new List<string>();
                        foreach (DataColumn column in table.Columns)
                        {
                            columnNames.Add(column.ColumnName);
                        }

                        foreach (DataRow row in table.Rows)
                        {
                            var rowValues = new List<Tuple<string, object>>();

                            for (int i = 0; i < row.ItemArray.Length; i++)
                            {
                                rowValues.Add(new Tuple<string, object>(columnNames[i], row.ItemArray[i]));
                                worksheet.Value.Add(rowValues);
                            }
                        }

                        workbook.Add(worksheet.Key, worksheet.Value);
                    }
                }
            }

            return workbook;
        }
    }
}
