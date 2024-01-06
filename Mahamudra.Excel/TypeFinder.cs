using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Mahamudra.Excel
{
    public class TypeFinder
    {
        private static readonly Dictionary<Type, CellValues> _matches = new Dictionary<Type, CellValues>()
        {
            { typeof(String) ,  CellValues.String},
            { typeof(Int64) ,  CellValues.Number},
            { typeof(Double) ,  CellValues.Number},
            { typeof(Int32) ,  CellValues.Number},
            { typeof(Int16) ,  CellValues.Number},
            { typeof(decimal) ,  CellValues.Number},
            { typeof(byte) ,  CellValues.Number},
            { typeof(Boolean) ,  CellValues.Boolean},
            { typeof(DateTime) ,  CellValues.Date},
            { typeof(DateTimeOffset) ,  CellValues.Date}
        };

        public static  (CellValues, T, Type) Get<T>(T value)
        {
            var type = value.GetType(); 
            var check = _matches.TryGetValue(type, out CellValues cellValues);
            if (check)
                return (cellValues, value, type); 
            return (CellValues.String, value, type);
        } 
    }
}