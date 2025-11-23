using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Mahamudra.Excel.Abstractions;
using Mahamudra.Excel.Domain;

namespace Mahamudra.Excel.Infrastructure
{
    /// <summary>
    /// Writes DataSet to Excel files using OpenXML.
    /// </summary>
    public sealed class ExcelWriter : IExcelWriter
    {
        /// <inheritdoc/>
        public MemoryStream Write(DataSet dataSet)
        {
            if (dataSet == null)
                throw new ArgumentNullException(nameof(dataSet));

            var memoryStream = new MemoryStream();
            using (var workbook = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart!.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                var maxDigitFont = 11;
                StylesheetBuilder.AddStyleSheet(workbook);

                uint sheetId = 1;

                foreach (DataTable table in dataSet.Tables)
                {
                    var numbersOfChars = new Dictionary<int, int?>();
                    for (var j = 0; j < table.Columns.Count; j++)
                    {
                        var len = table.Columns[j].Caption.Length;
                        numbersOfChars.TryGetValue(j, out var value);
                        if (value == null)
                            numbersOfChars.TryAdd(j, len);
                        else if (value < len)
                            numbersOfChars[j] = len;
                    }

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets!.Elements<Sheet>().Any())
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId!.Value).Max() + 1;

                    var sheet = new Sheet { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    var headerRow = new Row();
                    var columns = new Dictionary<string, XCellStyle>();
                    var clmns = new Columns();
                    var index = 0;

                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName, (XCellStyle)column.ExtendedProperties["Style"]!);
                        var widthPixels = Math.Truncate(((256 * numbersOfChars[index]!.Value + Math.Truncate(128f / maxDigitFont)) / 256f) * maxDigitFont);
                        var width = Math.Truncate(((widthPixels - 5f) / maxDigitFont * 100f + 0.5f) / 100f);
                        var cln = new Column
                        {
                            Min = Convert.ToUInt32(index + 1),
                            Max = Convert.ToUInt32(index + 1),
                            Width = width + 5,
                            CustomWidth = true,
                            Style = Convert.ToUInt32(0),
                        };
                        clmns.Append(cln);
                        index++;
                    }

                    var sheetdata = sheetPart.Worksheet.GetFirstChild<SheetData>();
                    sheetPart.Worksheet.InsertBefore(clmns, sheetdata);

                    foreach (DataColumn column in table.Columns)
                    {
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.Caption),
                            StyleIndex = Convert.ToUInt32(XCellStyle.Header),
                        };
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row();
                        foreach (var col in columns)
                        {
                            var (cellType, value, type) = CellTypeMapper.GetCellType(dsrow[col.Key]!);
                            var cell = new Cell
                            {
                                DataType = cellType,
                                CellValue = type == typeof(long) ? new CellValue(Convert.ToDecimal(value)) : new CellValue((dynamic)value),
                                StyleIndex = Convert.ToUInt32(col.Value),
                            };
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                }
            }
            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }
    }
}
