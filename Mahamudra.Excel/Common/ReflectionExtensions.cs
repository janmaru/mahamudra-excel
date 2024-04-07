using System.Collections.Generic;
using System.Data;
using System;
using System.Linq;

namespace Mahamudra.Excel.Common
{
    public static class ReflectionExtensions
    {
        internal static List<HeaderAttribute> GetHeaders<T>()
        {
            var list = new List<HeaderAttribute>();

            var myPropertyInfo = typeof(T).GetProperties();
            for (var i = 0; i < myPropertyInfo.Length; i++)
            {
                var customAttribute = (HeaderAttribute)Attribute.GetCustomAttribute(myPropertyInfo[i], typeof(HeaderAttribute));
                if (customAttribute != null)
                {
                    customAttribute.Type = myPropertyInfo[i].PropertyType;
                    customAttribute.Name = myPropertyInfo[i].Name;
                    list.Add(customAttribute);
                }
            }
            return list.OrderBy(x => x.Order).ToList();
        }

        internal static (DataTable, List<HeaderAttribute>, Dictionary<int, int?>) ToTable<T>()
        {
            var headers = GetHeaders<T>();
            var table = new DataTable(nameof(T));
            var numbersOfChars = new Dictionary<int, int?>();
            var cindex = 0;
            foreach (var hh in headers)
            {
                var column = new DataColumn
                {
                    DataType = Nullable.GetUnderlyingType(hh.Type!) ?? hh.Type,
                    ColumnName = hh.Name,
                    Caption = hh.Caption,
                    ReadOnly = hh.ReadOnly
                };
                column.ExtendedProperties.Add("Style", hh.Style);
                table.Columns.Add(column);

                // computing length chars
                var len = column.Caption.Length;
                numbersOfChars.TryGetValue(cindex, out var value);
                if (value == null)
                    numbersOfChars.TryAdd(cindex, len);
                else if (value < len)
                    numbersOfChars[cindex] = len;
                //
                cindex++;
            }
            return (table, headers, numbersOfChars);
        }

        internal static (DataTable, Dictionary<int, int?>) Fill<T>(this IEnumerable<T> data)
        { 
            var (table, headers, numbersOfChars) = ToTable<T>();
            DataRow row;
            foreach (var d in data)
            {
                row = table.NewRow();
                var cindex = 0;
                foreach (var hh in headers)
                {
                    row[hh.Name] = d.GetType().GetProperty(hh.Name).GetValue(d, null);

                    //
                    var len = row[hh.Name].ToString().Length;
                    numbersOfChars.TryGetValue(cindex, out var value);
                    if (value == null)
                        numbersOfChars.TryAdd(cindex, len);
                    else if (value < len)
                        numbersOfChars[cindex] = len;
                    //
                    cindex++;
                }
                table.Rows.Add(row);
            } 
            return (table, numbersOfChars);
        }

        internal static (DataSet, List<Dictionary<int, int?>>) Fill<T>(this ICollection<IEnumerable<T>> data)
        {
            var ds = new DataSet();
            var nc = new List<Dictionary<int, int?>>();

            foreach (var d in data)
            {
                var (t, c) = d.Fill<T>();
                ds.Tables.Add(t);
                nc.Add(c);
            } 
            return (ds, nc);
        }

        internal static DataSet FillOneSheet<T>(this IEnumerable<T> data)
        {
            var ds = new DataSet();
            var (table, headers, _) = ToTable<T>();
            DataRow row;
            foreach (var d in data)
            {
                row = table.NewRow();
                foreach (var hh in headers)
                    row[hh.Name!] = d!.GetType().GetProperty(hh.Name!)!.GetValue(d, null);

                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
            return ds;
        }
    }
}