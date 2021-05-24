using ExcelReader.Data.Loader;
using ExcelReader.Data.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReader.Tester
{
    class Program
    {
        private static readonly List<string> filePaths = new List<string>() {
                @""
        };

        static void Main(string[] args)
        {
            TestV1();
            Console.ReadKey();
        }

        private static void TestV1()
        {
            var server = new ExcelLoaderV1().LoadAll(filePaths);

            string output = string.Join("\n\n\n", server.Select(database =>
                $"Datenbank: {database.Key}\n\n"
                + string.Join("\n\n", database.Value.Select(table =>
                    $"{table.Key}:"
                    + "\n" + string.Join(" | ", table.Value.Item1.Select(col => col.Replace("\n", " ")))
                    + "\n" + string.Join("\n", table.Value.Item2.Select(row =>
                        string.Join(" | ", row.Select(cell => (!string.IsNullOrWhiteSpace(cell.Value?.ToString()) ? cell.Value.ToString().Replace("\n", " ") : string.Empty)))
                    ))
                ))
            ));

            Console.Write(output);
        }

        private static void TestV2()
        {
            List<Database> server = new ExcelLoaderV2().LoadAll(filePaths);

            string output = string.Join("\n\n\n", server.Select(database =>
                $"Datenbank: {database.Name}\n\n"
                + string.Join("\n\n", database.Tables.Select(table =>
                    $"{table.Name}:"
                    + "\n" + string.Join(" | ", table.Columns.Select(col => col.Name.Replace("\n", " ")))
                    + "\n" + string.Join("\n", table.Rows.Select(row =>
                        string.Join(" | ", row.Cells.Select(cell => (!string.IsNullOrWhiteSpace(cell.Value?.ToString()) ? cell.Value.ToString().Replace("\n", " ") : string.Empty)))
                    ))
                ))
            ));

            Console.Write(output);
        }
    }
}
