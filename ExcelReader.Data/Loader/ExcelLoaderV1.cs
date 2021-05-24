using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelReader.Data.Loader
{
    public class ExcelLoaderV1
    {
        public List<KeyValuePair<string, Dictionary<string, Tuple<List<string>, List<Dictionary<string, object>>>>>> LoadAll(List<string> filePaths)
        {
            var server = new List<KeyValuePair<string, Dictionary<string, Tuple<List<string>, List<Dictionary<string, object>>>>>>();

            foreach(string filePath in filePaths)
            {
                server.Add(Load(filePath));
            }

            return server;
        }
        
        // Database<DatabaseName, Tables<TableName, <Columns, Rows<ColumnName, Value>>>>
        public KeyValuePair<string, Dictionary<string, Tuple<List<string>, List<Dictionary<string, object>>>>> Load(string filePath)
        {
            var database = new KeyValuePair<string, Dictionary<string, Tuple<List<string>, List<Dictionary<string, object>>>>>(
                filePath.Split('\\')[filePath.Split('\\').Length - 1].Split('.')[0],
                new Dictionary<string, Tuple<List<string>, List<Dictionary<string, object>>>>()
            );

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

                    foreach (DataTable dataTable in result.Tables)
                    {
                        var table = new KeyValuePair<string, Tuple<List<string>, List<Dictionary<string, object>>>>(
                            dataTable.TableName,
                            new Tuple<List<string>, List<Dictionary<string, object>>>(
                                new List<string>(),
                                new List<Dictionary<string, object>>()
                            )
                        );

                        var columns = new List<string>();
                        foreach (DataColumn dataColumn in dataTable.Columns)
                        {
                            if (dataColumn.ColumnName.StartsWith("Column"))
                                break;

                            columns.Add(dataColumn.ColumnName);
                        }
                        table.Value.Item1.AddRange(columns);

                        foreach (DataRow dataRow in dataTable.Rows)
                        {
                            var row = new Dictionary<string, object>();
                            for (int i = 0; i < columns.Count; i++)
                            {
                                row.Add(columns[i], dataRow.ItemArray[i]);
                            }
                            table.Value.Item2.Add(row);
                        }

                        database.Value.Add(table.Key, table.Value);
                    }
                }
            }

            return database;
        }
    }
}
