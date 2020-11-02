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

        public ColumnTypeAttribute(Type? columnType, object? min = null, object? max = null) 
        {
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

                    if ((min as IComparable) == null)
                    {
                        throw new Exception($"ColumnTypeAttribute: The type of 'min' isn't comparable! [{min.GetType().Name}]");
                    }
                }

                if (max != null)
                {
                    if (max.GetType() != columnType)
                    {
                        throw new Exception($"ColumnTypeAttribute: No identical type for 'max' and 'column'! [{max.GetType().Name} vs. {columnType.Name}]");
                    }

                    if ((max as IComparable) == null)
                    {
                        throw new Exception($"ColumnTypeAttribute: The type of 'max' isn't comparable! [{max.GetType().Name}]");
                    }
                }
            }


            this.columnType = columnType; 
            this.min        = min; 
            this.max        = max; 
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
        All                 = TypeAndInterval
    }

    // 1 << 2,   // 4
    // 1 << 3,   // 8
}
