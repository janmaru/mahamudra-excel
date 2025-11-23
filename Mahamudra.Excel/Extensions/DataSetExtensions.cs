using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Mahamudra.Excel.Domain;
using Mahamudra.Excel.Infrastructure;

namespace Mahamudra.Excel.Extensions
{
    /// <summary>
    /// Extension methods for converting collections to DataSets and Excel files.
    /// </summary>
    public static class DataSetExtensions
    {
        /// <summary>
        /// Converts a DataSet to an Excel file stream.
        /// </summary>
        /// <param name="dataSet">The DataSet to convert.</param>
        /// <returns>A MemoryStream containing the Excel file.</returns>
        public static MemoryStream ToExcel(this DataSet dataSet)
        {
            var writer = new ExcelWriter();
            return writer.Write(dataSet);
        }

        /// <summary>
        /// Reads an Excel stream into a DataTable.
        /// </summary>
        /// <param name="stream">The memory stream containing the Excel file.</param>
        /// <returns>A DataTable with the Excel data.</returns>
        public static DataTable ReadExcel(this MemoryStream stream)
        {
            var reader = new ExcelReader();
            return reader.Read(stream);
        }

        /// <summary>
        /// Converts a collection to a DataSet with a single sheet.
        /// </summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="data">The collection to convert.</param>
        /// <param name="tableName">Optional name for the sheet/table.</param>
        /// <returns>A DataSet containing the data.</returns>
        public static DataSet FillOneSheet<T>(this IEnumerable<T> data, string? tableName = null)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            var dataSet = new DataSet();
            var (table, headers, _) = CreateTable<T>(tableName);

            foreach (var item in data)
            {
                var row = table.NewRow();
                foreach (var header in headers)
                {
                    var value = item!.GetType().GetProperty(header.Name!)!.GetValue(item, null);
                    row[header.Name!] = (value == null || (value is string s && string.IsNullOrEmpty(s)))
                        ? DBNull.Value
                        : value;
                }

                table.Rows.Add(row);
            }

            dataSet.Tables.Add(table);
            return dataSet;
        }

        /// <summary>
        /// Converts a collection to a tuple containing DataTable and column width information.
        /// </summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="data">The collection to convert.</param>
        /// <returns>A tuple of DataTable and column widths dictionary.</returns>
        internal static (DataTable, Dictionary<int, int?>) Fill<T>(this IEnumerable<T> data)
        {
            var (table, headers, numbersOfChars) = CreateTable<T>();

            foreach (var item in data)
            {
                var row = table.NewRow();
                var columnIndex = 0;

                foreach (var header in headers)
                {
                    row[header.Name!] = item!.GetType().GetProperty(header.Name!)!.GetValue(item, null);

                    var len = row[header.Name!]?.ToString()?.Length ?? 0;
                    numbersOfChars.TryGetValue(columnIndex, out var value);
                    if (value == null)
                        numbersOfChars.TryAdd(columnIndex, len);
                    else if (value < len)
                        numbersOfChars[columnIndex] = len;

                    columnIndex++;
                }

                table.Rows.Add(row);
            }

            return (table, numbersOfChars);
        }

        internal static (DataTable, List<HeaderAttribute>, Dictionary<int, int?>) CreateTable<T>(string? tableName = null)
        {
            var headers = GetHeaders<T>();
            var table = new DataTable(tableName ?? typeof(T).Name);
            var numbersOfChars = new Dictionary<int, int?>();
            var columnIndex = 0;

            foreach (var header in headers)
            {
                var column = new DataColumn
                {
                    DataType = Nullable.GetUnderlyingType(header.Type!) ?? header.Type,
                    ColumnName = header.Name,
                    Caption = header.Caption,
                    ReadOnly = header.ReadOnly
                };
                column.ExtendedProperties.Add("Style", header.Style);
                table.Columns.Add(column);

                var len = column.Caption.Length;
                numbersOfChars.TryGetValue(columnIndex, out var value);
                if (value == null)
                    numbersOfChars.TryAdd(columnIndex, len);
                else if (value < len)
                    numbersOfChars[columnIndex] = len;

                columnIndex++;
            }

            return (table, headers, numbersOfChars);
        }

        private static List<HeaderAttribute> GetHeaders<T>()
        {
            var list = new List<HeaderAttribute>();
            var properties = typeof(T).GetProperties();

            foreach (var property in properties)
            {
                var attribute = (HeaderAttribute?)Attribute.GetCustomAttribute(property, typeof(HeaderAttribute));
                if (attribute != null)
                {
                    attribute.Type = property.PropertyType;
                    attribute.Name = property.Name;
                    list.Add(attribute);
                }
            }

            return list.OrderBy(x => x.Order).ToList();
        }
    }
}
