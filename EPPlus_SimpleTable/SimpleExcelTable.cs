#nullable enable   

using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;

using OfficeOpenXml;                                                            // Install-Package EPPlus -Version 4.5.3.3      -- for excel; FREE (from version 5 it's commercial)

namespace EPPlus.SimpleTable
{
    public class SimpleExcelTable<T>  : IDisposable  where T : notnull, System.Enum 
    {
        public ExcelPackage     excel       { get; private set;}
        public ExcelWorksheet   worksheet   { get; private set;}

        public const string     extensionXLSX = ".XLSX";
        public const string     extensionXLS  = ".XLS";

        #region constructor

        public SimpleExcelTable(string excelName, string worksheetName, bool writeHeaderLine = true, bool updateHeaderLine = false) :
            this(GetExcelAndWorksheet(excelName, worksheetName), writeHeaderLine, updateHeaderLine)
        {            
        }

        public SimpleExcelTable(ExcelPackage excel, string worksheetName, bool writeHeaderLine = true, bool updateHeaderLine = false) :
            this (excel, GetWorksheet(excel, worksheetName), writeHeaderLine, updateHeaderLine)
        {            
        }

        public SimpleExcelTable((ExcelPackage excel, ExcelWorksheet worksheet) excelAndWorksheet, bool writeHeaderLine = true, bool updateHeaderLine = false) :
            this(excelAndWorksheet.excel, excelAndWorksheet.worksheet, writeHeaderLine, updateHeaderLine)
        {
        }
        
        public SimpleExcelTable(ExcelPackage excel, ExcelWorksheet worksheet, bool writeHeaderLine = true, bool updateHeaderLine = false)
        {
            if (! CheckEnum<T>())
            {
                throw new ArgumentException("Enum Type invalid! Enum's values must starts from one and increments by one; because enum value can index column of worksheet table as EPPlus's worksheet.Cells[] do it.");
            }

            this.excel     = excel;
            this.worksheet = worksheet;

            int rowCount = worksheet.Dimension.End.Row;
            int colCount = worksheet.Dimension.End.Column + 1;

            if ((rowCount == 0) && (colCount == 0))
            {   // Empty xls table
                if (writeHeaderLine || updateHeaderLine)
                {
                    WriteHead(true);
                }
            }
            else if (writeHeaderLine)
            {
                WriteHead(updateHeaderLine);
            }
        }
        #endregion

        #region constructor helper
        public static ExcelWorksheet GetWorksheet(ExcelPackage excel, string worksheetName)
        {
            if (String.IsNullOrWhiteSpace(worksheetName))
            {
                worksheetName = "SimpleExcelTableDefault";
            }

            foreach (var worksheet in excel.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    return worksheet;
                }
            }           

            return excel.Workbook.Worksheets.Add(worksheetName);
        }

        public static (ExcelPackage excel, ExcelWorksheet worksheet) GetExcelAndWorksheet(string excelName, string worksheetName)
        {
            // TODO

            string ext = (Path.GetExtension(excelName) ?? String.Empty).ToUpperInvariant();

            if ((ext != extensionXLS) && (ext != extensionXLSX))
            {
                Path.Combine(excelName, extensionXLSX);
            }

            var create = ! File.Exists(excelName);

            FileInfo        excelFile   = new FileInfo(excelName);
            ExcelPackage    excel       = create ? new ExcelPackage() : new ExcelPackage(excelFile);  
            ExcelWorksheet  worksheet   = GetWorksheet(excel, worksheetName);
            
            if (create)
            {  
                excel.SaveAs(excelFile);
            }
            
            return (excel, worksheet);
        }

        public static bool CheckEnum<TEnum>() where TEnum : notnull, System.Enum 
        {
            var values = Enum.GetValues(typeof(TEnum));

            if (values == null)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < values.Length; i++)
                {
                    int enumVal = (int)(values.GetValue(i));

                    if ((i + 1) != enumVal)
                    {
                        return false;
                    }
                }
            }

            return true;
        }
        #endregion

        #region excel read/write

        public void WriteHead(bool updateHeaderLine = false)
        {
            int rowCount = worksheet.Dimension.End.Row;
            int colCount = worksheet.Dimension.End.Column + 1;

            throw new NotImplementedException();
        }
        #endregion

        #region IDisposable

        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (! disposedValue)
            {
                if (disposing)
                {                    
                }

                excel.Save();    

                disposedValue = true;
            }
        }

        ~SimpleExcelTable()
        {   
            Dispose(disposing: false);
        }

        void IDisposable.Dispose()
        {   
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region TEST & DEMO

        [Conditional("DEBUG")]
        public static void Test1()
        { 
            var filename = Path.ChangeExtension(Path.GetTempFileName(), extensionXLSX);

            using var test = new SimpleExcelTable<TestAndDemoEnum>(filename, "Test1");


            // TODO
            // TODO

            lastTestFilename = filename;
        }

        #if DEBUG

        public static string? lastTestFilename { get; private set; }

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
        #endif

#endregion
    }
}
