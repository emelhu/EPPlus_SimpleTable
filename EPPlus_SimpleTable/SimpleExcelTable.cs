#nullable enable   

using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;

using OfficeOpenXml;                                                            // Install-Package EPPlus -Version 4.5.3.3      -- for excel; FREE (from version 5 it's commercial)
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;


// worksheet.Cells[2, 1, (szamlak.Count + 1), 1].Style.Numberformat.Format = "yyyy-MM-dd";

namespace EPPlus.SimpleTable
{
    public class SimpleExcelTable<T>  : IDisposable  where T : notnull, System.Enum 
    {
        public ExcelPackage     excel       { get; private set;}
        public ExcelWorksheet   worksheet   { get; private set;}

        public int rowCount     => worksheet.Dimension?.End?.Row    ?? 0;
        public int colCount     => worksheet.Dimension?.End?.Column ?? 0;

        public const string     extensionXLSX = ".XLSX";
        public const string     extensionXLS  = ".XLS";

        #region constructor

        public SimpleExcelTable(string excelName, string worksheetName, bool writeHeaderLine = true) :
            this(GetExcelAndWorksheet(excelName, worksheetName), writeHeaderLine)
        {            
        }

        public SimpleExcelTable(ExcelPackage excel, string worksheetName, bool writeHeaderLine = true) :
            this (excel, GetWorksheet(excel, worksheetName), writeHeaderLine)
        {            
        }

        public SimpleExcelTable((ExcelPackage excel, ExcelWorksheet worksheet) excelAndWorksheet, bool writeHeaderLine = true) :
            this(excelAndWorksheet.excel, excelAndWorksheet.worksheet, writeHeaderLine)
        {
        }
        
        public SimpleExcelTable(ExcelPackage excel, ExcelWorksheet worksheet, bool writeHeaderLine = true)
        {
            if (! CheckEnum<T>())
            {
                throw new ArgumentException("Enum Type invalid! Enum's values must starts from one and increments by one; because enum value can index column of worksheet table as EPPlus's worksheet.Cells[] do it.");
            }

            this.excel     = excel;
            this.worksheet = worksheet;

            if (( worksheet.Dimension == null) || ((rowCount == 0) && (colCount == 0)))
            {   // Empty xls table
                if (writeHeaderLine)
                {
                    WriteHeaderFirstLine();
                }
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

        /// <summary>
        /// Write header line into first line of excel table
        /// </summary>
        public void WriteHeaderFirstLine()
        {          
            var enumValues = Enum.GetValues(typeof(T));
            int maxColumn  = 1;   

            foreach (var enumValue in enumValues)
            {
                if (enumValue != null)
                {
                    var name = GetDisplayName((T)enumValue);

                    worksheet.SetValue(1, (int)enumValue, name);

                    if (maxColumn < (int)enumValue)
                    {
                        maxColumn = (int)enumValue;
                    }
                }
            }

            using (var range = worksheet.Cells[1, 1, 1, maxColumn])
            {
                range.Style.Font.Italic = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                range.Style.Font.Color.SetColor(Color.DarkBlue);
            }
        }
        #endregion

        #region Helper functions

        public string GetDisplayName(T enumValue)
        {
            string name = enumValue.ToString();

            string? dispName = enumValue.GetType()?
                 .GetMember(enumValue.ToString())?[0]?
                 .GetCustomAttribute<DisplayAttribute>()?
                 .Name;

            if (dispName != null)
            {
                return dispName;
            }

            return name;
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

    }
}
