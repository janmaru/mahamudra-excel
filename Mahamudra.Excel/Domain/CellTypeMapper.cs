using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace Mahamudra.Excel.Domain
{
    /// <summary>
    /// Maps .NET types to Excel cell value types.
    /// </summary>
    public static class CellTypeMapper
    {
        private static readonly IReadOnlyDictionary<Type, CellValues> _typeMap = new Dictionary<Type, CellValues>
        {
            { typeof(string), CellValues.String },
            { typeof(long), CellValues.Number },
            { typeof(double), CellValues.Number },
            { typeof(int), CellValues.Number },
            { typeof(short), CellValues.Number },
            { typeof(decimal), CellValues.Number },
            { typeof(byte), CellValues.Number },
            { typeof(bool), CellValues.Boolean },
            { typeof(DateTime), CellValues.Date },
            { typeof(DateTimeOffset), CellValues.Date }
        };

        /// <summary>
        /// Gets the Excel cell type for a given value.
        /// </summary>
        /// <typeparam name="T">The type of the value.</typeparam>
        /// <param name="value">The value to map.</param>
        /// <returns>A tuple containing the cell type, value, and .NET type.</returns>
        public static (CellValues CellType, T Value, Type Type) GetCellType<T>(T value)
        {
            if (value == null || value is DBNull)
                return (CellValues.String, default!, typeof(string));

            var type = value.GetType();
            if (_typeMap.TryGetValue(type, out var cellValues))
                return (cellValues, value, type);

            return (CellValues.String, default!, typeof(string));
        }
    }
}
