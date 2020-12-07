#nullable enable   

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

using OfficeOpenXml;                                                            // Install-Package EPPlus -Version 4.5.3.3      -- for excel; FREE (from version 5 it's commercial)
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;


// worksheet.Column(1).Style.Numberformat.Format  = "yyyy-mm-dd"; 


namespace EPPlus.SimpleTable
{
    public class SimpleExcelTable<T>  : IDisposable  where T : notnull, System.Enum 
    {
        public ExcelPackage     excel       { get; private set;}
        public ExcelWorksheet   worksheet   { get; private set;}

        public int rowCount     => worksheet.Dimension?.End?.Row    ?? 0;
        public int colCount     => worksheet.Dimension?.End?.Column ?? 0;

        public IEnumerable<string>                      workbookNumberformats       => excel.Workbook.Styles.NumberFormats.Select(i => i.Format);
        public IEnumerable<(int id, string format)>     workbookNumberformatsWithId => excel.Workbook.Styles.NumberFormats.Select(i => (i.NumFmtId, i.Format));

        public const string     extensionXLSX        = ".XLSX";
        public const string     extensionXLS         = ".XLS";
        public const string     defaultWorksheetName = "Default";

        /// <summary>
        /// Default values for set 'worksheet.Cells[].Style.Numberformat.Format' automatically.
        /// </summary>
        public static Dictionary<Type, string> numberFormatsForTypesDefault { get; private set; }

        /// <summary>
        /// Values for set 'worksheet.Cells[].Style.Numberformat.Format' automatically.
        /// This variable inherints content from numberFormatsDefault but you can modify it.
        /// </summary>
        public        Dictionary<Type, string> numberFormatsForTypes        { get; private set; }

        public static Appropriateness   defaultAppropriatenessDefault   = Appropriateness.None;
        public        Appropriateness   defaultAppropriateness          = defaultAppropriatenessDefault;

        #region constructor

        static SimpleExcelTable()
        {
            numberFormatsForTypesDefault = new Dictionary<Type, string>();

            numberFormatsForTypesDefault.Add(typeof(DateTime),  DateTimeFormatInfo.CurrentInfo.ShortDatePattern);
            numberFormatsForTypesDefault.Add(typeof(int),       "0");
            numberFormatsForTypesDefault.Add(typeof(double),    "#,##0.00");
            numberFormatsForTypesDefault.Add(typeof(decimal),   "#,##0.00");
            numberFormatsForTypesDefault.Add(typeof(string),    "@");            
        }

        public SimpleExcelTable(string excelName, string? worksheetName, bool writeUniformFormatting = true, params object[]? columnFormats) :
            this(GetExcelAndWorksheet(excelName, worksheetName), writeUniformFormatting, columnFormats)
        {            
        }

        public SimpleExcelTable(ExcelPackage excel, string? worksheetName, bool writeUniformFormatting = true, params object[]? columnFormats) :
            this (excel, GetWorksheet(excel, worksheetName), writeUniformFormatting, columnFormats)
        {            
        }

        public SimpleExcelTable((ExcelPackage excel, ExcelWorksheet worksheet) excelAndWorksheet, bool writeUniformFormatting = true, params object[]? columnFormats) :
            this(excelAndWorksheet.excel, excelAndWorksheet.worksheet, writeUniformFormatting, columnFormats)
        {
        }
        
        public SimpleExcelTable(ExcelPackage excel, ExcelWorksheet worksheet, bool writeUniformFormatting = true, params object[]? columnFormats)
        {
            if (! CheckEnum<T>())
            {
                throw new ArgumentException("Enum Type invalid! Enum's values must starts from one and increments by one; because enum value can index column of worksheet table as EPPlus's worksheet.Cells[] do it.");
            }

            numberFormatsForTypes = numberFormatsForTypesDefault.ToDictionary(entry => entry.Key, entry => entry.Value);

            // TODO: columnFormats parameter

            this.excel     = excel;
            this.worksheet = worksheet;

            if (( worksheet.Dimension == null) || ((rowCount == 0) && (colCount == 0)))
            {   // Empty xls table
                if (writeUniformFormatting)
                {
                    WriteUniformFormatting();
                }
            }            
        }
        #endregion

        #region constructor helper
        public static ExcelWorksheet GetWorksheet(ExcelPackage excel, string? worksheetName)
        {
            var worksheets = excel.Workbook.Worksheets.ToArray();

            if (String.IsNullOrWhiteSpace(worksheetName))
            {
                if (worksheets.Length == 0)
                {
                    worksheetName = defaultWorksheetName; 
                }
                else
                {
                    return worksheets[0];
                }              
            }

            foreach (var worksheet in worksheets)
            {
                if (worksheet.Name == worksheetName)
                {
                    return worksheet;
                }
            }           

            return excel.Workbook.Worksheets.Add(worksheetName);
        }

        public static (ExcelPackage excel, ExcelWorksheet worksheet) GetExcelAndWorksheet(string excelName, string? worksheetName)
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

        #region excel table read/write

        /// <summary>
        /// Write header line into first line of excel table and column formatting (if defined)
        /// </summary>
        public void WriteUniformFormatting()
        {          
            var enumValues = Enum.GetValues(typeof(T));
            int maxColumn  = 1;   

            foreach (var enumValue in enumValues)
            {
                if (enumValue != null)
                {
                    var name = GetColumnName((T)enumValue);

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

                range.Style.Numberformat.Format = "@";
                range.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Blue);
            }
        }
        #endregion

        #region indexer

        public object this[int row, T col, Appropriateness checkAppropriateness = Appropriateness.Default]
        {
            get { return worksheet.GetValue(row, Convert.ToInt32(col)); }
            set { CheckAppropriateness(col, value, checkAppropriateness);
                  worksheet.SetValue(row, Convert.ToInt32(col), value); }
        }

        public object this[int row, T col, string styleNumberFormat, Appropriateness checkAppropriateness = Appropriateness.Default]
        {
            get { return worksheet.GetValue(row, Convert.ToInt32(col)); }
            set { CheckAppropriateness(col, value, checkAppropriateness);
                  worksheet.SetValue(row, Convert.ToInt32(col), value); 
                  worksheet.Cells[row, Convert.ToInt32(col)].Style.Numberformat.Format = styleNumberFormat;}
        }

        public object this[int row, T col, int styleNumberFormatId, Appropriateness checkAppropriateness = Appropriateness.Default]
        {
            get { return worksheet.GetValue(row, Convert.ToInt32(col)); }
            set { CheckAppropriateness(col, value, checkAppropriateness);
                  worksheet.SetValue(row, Convert.ToInt32(col), value); 
                  worksheet.Cells[row, Convert.ToInt32(col)].Style.Numberformat.Format = GetNumberformat(styleNumberFormatId) ?? "General";}
        }

        public object this[int row, T col, bool uniformFormat, Appropriateness checkAppropriateness = Appropriateness.Default]
        {
            get { return worksheet.GetValue(row, Convert.ToInt32(col)); }

            set
            {     
                CheckAppropriateness(col, value, checkAppropriateness);
                worksheet.SetValue(row, Convert.ToInt32(col), value);

                if (uniformFormat)
                {
                    string? format  = GetColumnNumberformat(col);

                    if (format == null)
                    {
                        format = GetNumberformat(value.GetType());
                    }
                                        
                    worksheet.Cells[row, Convert.ToInt32(col)].Style.Numberformat.Format = format ?? "General";
                }
            }
        }

        public object this[int row, T col, Type setNumberFormatByType, Appropriateness checkAppropriateness = Appropriateness.Default]
        {
            get { return worksheet.GetValue(row, Convert.ToInt32(col)); }
            set { CheckAppropriateness(col, value, checkAppropriateness);
                  worksheet.SetValue(row, Convert.ToInt32(col), value); 
                  worksheet.Cells[row, Convert.ToInt32(col)].Style.Numberformat.Format = GetNumberformat(setNumberFormatByType) ?? "General";}
        }

        #endregion

        #region Helper functions

        public string GetColumnName(T enumValue)
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

        public Type? GetColumnType(T enumValue)
        {
            Type? type = enumValue.GetType()?
                  .GetMember(enumValue.ToString())?[0]?
                  .GetCustomAttribute<ColumnTypeAttribute>()?
                  .columnType;

            return type;
        }

        public (Type? type, object? min, object? max, int? minLen, int? maxLen)? GetColumnCheckData(T enumValue)
        {
            var attr = enumValue.GetType()?
                  .GetMember(enumValue.ToString())?[0]?
                  .GetCustomAttribute<ColumnTypeAttribute>();

            if (attr != null)
            { 
                return (attr.columnType, attr.min, attr.max, attr.minLen, attr.maxLen);
            }

            return null;
        }

        public string? GetColumnNumberformat(T enumValue)
        {
            return enumValue.GetType()?
                   .GetMember(enumValue.ToString())?[0]?
                   .GetCustomAttribute<ColumnNumberformatAttribute>()?
                   .columnNumberformat;
        }

        public string? GetNumberformat(int id)
        {
            return excel.Workbook.Styles.NumberFormats.Where(i => i.NumFmtId == id).Select(i => i.Format).FirstOrDefault();
        }

        public string? GetNumberformat(Type type)
        {
            return numberFormatsForTypes.Where(i => i.Key == type).Select(i => i.Value).FirstOrDefault();
        }

        public void CheckAppropriateness(T col, object storeValue, Appropriateness checkAppropriateness)
        {
            if (checkAppropriateness == Appropriateness.Default)
            {
                checkAppropriateness = defaultAppropriateness;
            }
            

            if (checkAppropriateness != Appropriateness.None)
            {
                var checkData = GetColumnCheckData(col);

                if (checkData != null)
                {
                    if ((checkAppropriateness & Appropriateness.Type) != 0)                                                             // you want type check
                    {
                        if ((checkData.Value.type != null) && (checkData.Value.type != storeValue.GetType()))
                        {
                            throw new Exception($"The type of value for store to excel worksheet's column '{col.ToString()}' is not identical then defined by enum! [{checkData.Value.type.Name} vs. {storeValue.GetType().Name}]");
                        }                   
                    }

                    if ((checkAppropriateness & Appropriateness.Interval) != 0)                                                         // you want interval check
                    {
                        if (checkData.Value.min != null) 
                        {
                            if (checkData.Value.min.GetType() != storeValue.GetType()) 
                            {
                                throw new Exception($"The type of value for store to excel worksheet's column '{col.ToString()}' is not identical then defined 'min' value! [{checkData.Value.min.GetType().Name} vs. {storeValue.GetType().Name}]");
                            }

                            var comparable = (IComparable)checkData.Value.min as IComparable;

                            if (comparable.CompareTo(storeValue) < 0) 
                            {
                                throw new Exception($"The value for store to excel worksheet's column '{col.ToString()}' is less then defined 'min' value! [{checkData.Value.min.ToString()} vs. {storeValue.ToString()}]");
                            }
                        }

                        if (checkData.Value.max != null) 
                        {
                            if (checkData.Value.max.GetType() != storeValue.GetType()) 
                            {
                                throw new Exception($"The type of value for store to excel worksheet's column '{col.ToString()}' is not identical then defined 'max' value! [{checkData.Value.max.GetType().Name} vs. {storeValue.GetType().Name}]");
                            }

                            var comparable = (IComparable)checkData.Value.max as IComparable;

                            if (comparable.CompareTo(storeValue) > 0) 
                            {
                                throw new Exception($"The value for store to excel worksheet's column '{col.ToString()}' is more then defined 'max' value! [{checkData.Value.max.ToString()} vs. {storeValue.ToString()}]");
                            }
                        }

                        if (checkData.Value.minLen != null)
                        {
                            var len = storeValue.ToString().Length;

                            if (len < checkData.Value.minLen) 
                            {
                                throw new Exception($"The length of value for store to excel worksheet's column '{col.ToString()}' is shorter then defined 'minLen' value! [{checkData.Value.minLen} > {len}] [{storeValue.ToString()}]");
                            }
                        }

                        if (checkData.Value.maxLen != null)
                        {
                            var len = storeValue.ToString().Length;

                            if (len > checkData.Value.maxLen) 
                            {
                                throw new Exception($"The length of value for store to excel worksheet's column '{col.ToString()}' is longer then defined 'maxLen' value! [{checkData.Value.maxLen} < {len}] [{storeValue.ToString()}]");
                            }
                        }
                    }
                }
            }
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

#region other info
/*
build in 'Workbook.Styles.NumberFormats':                        // EPPlus -Version 4.5.3.3
 'General'                  ; id: 0             // auto, apply the default number format
 '0'                        ; id: 1             // Digit placeholder
 '0.00'                     ; id: 2             // Digit placeholder, two decimals
 '#,##0'                    ; id: 3             // Digit placeholder, does not display extra zeros, thousands separator
 '#,##0.00'                 ; id: 4             // Digit placeholder, does not display extra zeros, thousands separator, two decimals
 '0%'                       ; id: 9             // Percentage
 '0.00%'                    ; id: 10            // Percentage
 '0.00E+00'                 ; id: 11            // Scientific format
 '# ?/?'                    ; id: 12
 '# ??/??'                  ; id: 13
 'mm-dd-yy'                 ; id: 14            // Date format
 'd-mmm-yy'                 ; id: 15            // Date format
 'd-mmm'                    ; id: 16            // Date format
 'mmm-yy'                   ; id: 17            // Date format
 'h:mm AM/PM'               ; id: 18            // Time format
 'h:mm:ss AM/PM'            ; id: 19            // Time format
 'h:mm'                     ; id: 20            // Time format
 'h:mm:ss'                  ; id: 21            // Time format
 'm/d/yy h:mm'              ; id: 22            // Date format with time
 '#,##0 ;(#,##0)'           ; id: 37            // Different view of positive and negative numbers
 '#,##0 ;[Red](#,##0)'      ; id: 38            // Different view of positive and negative numbers
 '#,##0.00;(#,##0.00)'      ; id: 39            // Different view of positive and negative numbers
 '#,##0.00;[Red](#,##0.00)' ; id: 40            // Different view of positive and negative numbers
 'mm:ss'                    ; id: 45            // Time format
 '[h]:mm:ss'                ; id: 46            // Time format
 'mmss.0'                   ; id: 47            // Time format
 '##0.0'                    ; id: 48            // Digit placeholder, one decimals
 '@'                        ; id: 49            // Text placeholder
*/

/*
- Each format that you create can have up to three sections for numbers and a fourth section for text: <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>
- To set the color for any section in the custom format, type the name of the color in brackets in the section. [BLUE]#,##0;[RED]#,##0   (positive numbers blue and negative numbers red)
- conditional statements available (see link above): [>100][GREEN]#,##0;[<=-100][YELLOW]#,##0;[CYAN]#,##0
*/

// useful links:
// https://docs.microsoft.com/en-us/office/troubleshoot/excel/format-cells-settings
// https://support.microsoft.com/en-us/help/81518/using-a-custom-number-format-to-display-leading-zeros

#endregion