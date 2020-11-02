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

            DispBuiltInNumberFormats(filename);
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

        public static void DispBuiltInNumberFormats(string filename)
        {
            using var test = new SimpleExcelTable<TestAndDemoEnum>(filename, worksheetName);

            Console.WriteLine("Display all hard-coded codes and keys of build in 'Workbook.Styles.NumberFormats':");
         
            foreach (var format in test.excel.Workbook.Styles.NumberFormats)
            {   // https://stackoverflow.com/questions/9859610/how-to-set-column-type-when-using-epplus
                Console.WriteLine($" Format built in: {format.BuildIn}; string: '{format.Format}'; id: {format.NumFmtId}");
            }
        }

        /// <summary>
        /// List of HEAD columns for 'Test & Demo'
        /// Values must starts from one and increments by one; because enum value can index column of worksheet table as EPPlus's worksheet.Cells[] do it.
        /// </summary>
        public enum TestAndDemoEnum
        {
            [ColumnType(typeof(int))]
            FIRST = 1,
            Second,
            [Display(Name = "Third column")]
            THIRD,
            [Display(Name = "Fourth column")]
            [ColumnType(typeof(double), 1.0, 1000)]                                                 // !ERR! --> 1000.0
            [ColumnNumberformat("#,##0.0000 thousand")]
            fourth
        }
    }
}
