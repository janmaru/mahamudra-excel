using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Mahamudra.Excel.Common
{
    public class TypeFinder
    {
        private static readonly Dictionary<Type, CellValues> _matches = new Dictionary<Type, CellValues>()
        {
            { typeof(string) ,  CellValues.String},
            { typeof(long) ,  CellValues.Number},
            { typeof(double) ,  CellValues.Number},
            { typeof(int) ,  CellValues.Number},
            { typeof(short) ,  CellValues.Number},
            { typeof(decimal) ,  CellValues.Number},
            { typeof(byte) ,  CellValues.Number},
            { typeof(bool) ,  CellValues.Boolean},
            { typeof(DateTime) ,  CellValues.Date},
            { typeof(DateTimeOffset) ,  CellValues.Date}
        };

        public static (CellValues, T, Type) Get<T>(T value)
        {
            var type = value.GetType();
            var check = _matches.TryGetValue(type, out var cellValues);
            if (check)
                return (cellValues, value, type);
            return (CellValues.String, value, type);
        }
    }
}