using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Mahamudra.Excel.Abstractions;

namespace Mahamudra.Excel.Infrastructure
{
    /// <summary>
    /// Reads Excel files into DataTables using OpenXML.
    /// </summary>
    public sealed class ExcelReader : IExcelReader
    {
        /// <inheritdoc/>
        public DataTable Read(MemoryStream stream)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));

            var table = new DataTable();
            using var spreadSheetDocument = SpreadsheetDocument.Open(stream, true);

            var workbookPart = spreadSheetDocument.WorkbookPart;
            var sheets = spreadSheetDocument.WorkbookPart!.Workbook.GetFirstChild<Sheets>()!.Elements<Sheet>();
            var relationshipId = sheets.First().Id!.Value!;
            var worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
            var workSheet = worksheetPart.Worksheet;
            var sheetData = workSheet.GetFirstChild<SheetData>();
            var rows = sheetData!.Descendants<Row>();

            foreach (var cell in rows.ElementAt(0).Cast<Cell>())
                table.Columns.Add(GetCellValue(spreadSheetDocument, cell));

            foreach (var row in rows)
            {
                var tempRow = table.NewRow();
                for (var i = 0; i < row.Descendants<Cell>().Count(); i++)
                    tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                table.Rows.Add(tempRow);
            }

            table.Rows.RemoveAt(0);
            return table;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var stringTablePart = document.WorkbookPart!.SharedStringTablePart;
            var value = cell.CellValue?.InnerXml ?? string.Empty;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && stringTablePart != null)
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;

            return value;
        }
    }
}
