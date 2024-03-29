﻿using System.Collections.Generic;
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

        internal static (DataTable, List<HeaderAttribute>) ToTable<T>()
        {
            var headers = GetHeaders<T>();
            var table = new DataTable(nameof(T));
            foreach (var hh in headers)
            {
                var column = new DataColumn
                {
                    DataType = hh.Type,
                    ColumnName = hh.Name,
                    Caption = hh.Caption,
                    ReadOnly = hh.ReadOnly
                };
                column.ExtendedProperties.Add("Style", hh.Style);
                table.Columns.Add(column);
            }
            return (table, headers);
        }

        internal static DataSet FillOneSheet<T>(this IEnumerable<T> data)
        {
            var ds = new DataSet();
            var (table, headers) = ToTable<T>();
            DataRow row;
            foreach (var d in data)
            {
                row = table.NewRow();
                foreach (var hh in headers)
                    row[hh.Name] = d.GetType().GetProperty(hh.Name).GetValue(d, null);

                table.Rows.Add(row);
            }
            ds.Tables.Add(table);
            return ds;
        }

        public static DataSet FillWorkbook<T>(this IEnumerable<DataTable> dataTables)
        {
            var ds = new DataSet();
            foreach (var d in dataTables)
                ds.Tables.Add(d);
            return ds;
        }
    }
}