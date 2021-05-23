using ExcelReader.Data.Loader;
using ExcelReader.Data.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReader.Tester
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelV1 = new ExcelLoaderV1().LoadAll(new List<string>() {
                @""
            });

            List<Workbook> excelV2 = new ExcelLoaderV2().LoadAll(new List<string>() {
                @""
            });

            Console.ReadKey();
        }
    }
}
