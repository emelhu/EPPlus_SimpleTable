using System;
using System.IO;

namespace Test
{
    using System.ComponentModel.DataAnnotations;

    using EPPlus.SimpleTable;

    class Program
    {
        const string worksheetName = "Test1";

        static void Main(string[] args)
        {
            Console.WriteLine("*** EPPlus.SimpleTable TEST ***\n");

            var filename = Path.Combine(Path.GetTempPath(), "SimpleExcelTable---Test---" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + SimpleExcelTable<TestAndDemoEnum>.extensionXLSX);

            Test1(filename);

            DispResult(filename);
        }

        public static void Test1(string filename)
        {             
            using var test = new SimpleExcelTable<TestAndDemoEnum>(filename, worksheetName);
        }

        public static void DispResult(string filename)
        {
            using var test = new SimpleExcelTable<TestAndDemoEnum>(filename, worksheetName);

            int rows = test.rowCount;
            int cols = test.colCount;

            Console.WriteLine($"The reuslt file is: {filename}");
            Console.WriteLine($" count of rows is {rows}, count of columns is {cols}");
        }

        /// <summary>
        /// List of HEAD columns for 'Test & Demo'
        /// Values must starts from one and increments by one; because enum value can index column of worksheet table as EPPlus's worksheet.Cells[] do it.
        /// </summary>
        public enum TestAndDemoEnum
        {
            FIRST = 1,
            Second,
            [Display(Name = "Third column")]
            THIRD
        }
    }
}
