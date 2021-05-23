using ExcelReader.Data.Loader;
using System;
using System.Collections.Generic;

namespace ExcelReader.Tester
{
    class Program
    {
        static void Main(string[] args)
        {
            var excel = new ExcelLoader().LoadAll(new List<string>() {
                @""
            });

            Console.ReadKey();
        }
    }
}
