#nullable enable   

using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.SimpleTable
{
    public class ColumnTypeAttribute : Attribute
    {
        public Type?    columnType;
        public object?  min;
        public object?  max;
        public int?     minLen;
        public int?     maxLen;

        private const int NullIntReplacement = int.MinValue;

        public ColumnTypeAttribute(Type? columnType, object? min = null, object? max = null, int minLen = NullIntReplacement, int maxLen = NullIntReplacement) 
        {
            this.columnType = columnType; 
            this.min        = min; 
            this.max        = max; 
            this.minLen     = minLen; 
            this.maxLen     = maxLen; 

            if (minLen == NullIntReplacement)
            {
                this.minLen = null;
            }

            if (maxLen == NullIntReplacement)
            {
                this.maxLen = null;
            }


            if (columnType == null)
            {
                if ((min != null) && (max != null))
                {
                    if (min.GetType() != max.GetType())
                    {
                        throw new Exception($"ColumnTypeAttribute: No identical type for 'min' and 'max'! [{min.GetType().Name} vs. {max.GetType().Name}]");
                    }
                }                
            }
            else
            {
                if (min != null)
                {
                    if (min.GetType() != columnType)
                    {
                        throw new Exception($"ColumnTypeAttribute: No identical type for 'min' and 'column'! [{min.GetType().Name} vs. {columnType.Name}]");
                    }
                }

                if (max != null)
                {
                    if (max.GetType() != columnType)
                    {
                        throw new Exception($"ColumnTypeAttribute: No identical type for 'max' and 'column'! [{max.GetType().Name} vs. {columnType.Name}]");
                    }
                }
            }

            if ((min != null) && ((min as IComparable) == null))
            {
                throw new Exception($"ColumnTypeAttribute: The type of 'min' isn't comparable! [{min.GetType().Name}]");
            }

            if ((max != null) && ((max as IComparable) == null))
            {
                throw new Exception($"ColumnTypeAttribute: The type of 'max' isn't comparable! [{max.GetType().Name}]");
            }


            const int stringLimit = 32767;                                                                                              // https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3?ocmsassetid=hp010073849&correlationid=bbfa300d-224f-47a5-bf71-7f71b9ae0761&ui=en-us&rs=en-us&ad=us

            if (this.minLen != null)
            {
                if (minLen < 0)
                {
                    throw new Exception($"ColumnTypeAttribute: 'minLen' less then zero! [{minLen}]");
                }

                if (minLen > stringLimit)                                                                                                                 
                {
                    throw new Exception($"ColumnTypeAttribute: 'minLen' more then limit! [{minLen}/{stringLimit}]");
                }
            }

            if (this.maxLen != null)
            {
                if (maxLen < 0)
                {
                    throw new Exception($"ColumnTypeAttribute: 'maxLen' less then zero! [{maxLen}]");
                }

                if (maxLen > stringLimit)                                                                                                                 
                {
                    throw new Exception($"ColumnTypeAttribute: 'maxLen' more then limit! [{maxLen}/{stringLimit}]");
                }
            }

            if ((this.minLen != null) && (this.maxLen != null))
            {
                if (minLen > maxLen)
                {
                    throw new Exception($"ColumnTypeAttribute: 'minLen' more then 'maxLen'! [{minLen}/{maxLen}]");
                }
            }            
        }
    }

    public class ColumnNumberformatAttribute : Attribute
    {
        public string columnNumberformat;

        public ColumnNumberformatAttribute(string columnNumberformat) { this.columnNumberformat = columnNumberformat; }
    }

    //

    [Flags]
    public enum Appropriateness
    {
        None                = 0,
        Type                = 1 << 0,   // 1
        Interval            = 1 << 1,   // 2
        TypeAndInterval     = Type | Interval,
        All                 = TypeAndInterval,
        Default             = -1
    }

    // 1 << 2,   // 4
    // 1 << 3,   // 8
}
