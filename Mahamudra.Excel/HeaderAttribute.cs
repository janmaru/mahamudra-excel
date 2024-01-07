using System;
using Mahamudra.Excel.Common;

namespace Mahamudra.Excel
{
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public class HeaderAttribute : Attribute
    { 
        public HeaderAttribute(string caption, bool readOnly = false)
        {
            this.Caption = caption; 
            this.ReadOnly = readOnly;
        }

        public string Caption { get; internal set; }
        public string Name { get; set; }
        public bool ReadOnly { get; internal set; } 
        public Type Type { get; set; }
        public Int16 Order { get; set; } = 0;
        public XCellStyle Style { get; set; } = XCellStyle.None;
    }
}
