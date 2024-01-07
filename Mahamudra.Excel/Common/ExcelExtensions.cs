using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet; 

namespace Mahamudra.Excel.Common
{
    public static class ExcelExtensions
    { 

        internal static WorkbookStylesPart AddStyleSheet(this SpreadsheetDocument spreadsheet)
        {
            var stylesheet = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            var workbookstylesheet = new Stylesheet();

            // <Fonts>
            var font0 = new Font();            // Default font
            var fonts = new Fonts();          // <APPENDING Fonts>
            fonts.Append(font0);

            // <Fills>
            var fill0 = new Fill();            // Default fill
            var fills = new Fills();          // <APPENDING Fills>
            fills.Append(fill0);

            // <Borders>
            var border0 = new Border();      // Defualt border
            var borders = new Borders();    // <APPENDING Borders>
            borders.Append(border0);

            // <CellFormats>
            var cellformatHeader = new CellFormat()   // Default style : Mandatory
            {
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                NumberFormatId = 0, 
                FormatId = 0
            };
            var cellformatRow = new CellFormat(new Alignment() { WrapText = true });
            // Style with textwrap set

            // <APPENDING CellFormats>
            var cellformats = new CellFormats();
            cellformats.Append(cellformatHeader);
            cellformats.Append(cellformatRow);

            // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
            workbookstylesheet.Append(fonts);
            workbookstylesheet.Append(fills);
            workbookstylesheet.Append(borders);
            workbookstylesheet.Append(cellformats);

            // Finalize
            stylesheet.Stylesheet = workbookstylesheet;
            stylesheet.Stylesheet.Save();

            return stylesheet;
        }

        public static bool SheetExist(this SpreadsheetDocument doc, string sheetName)
        { 
            if (doc == null) throw new ArgumentNullException("SpreadsheetDocument");
            if (doc.WorkbookPart == null) throw new ArgumentNullException("WorkbookPart");
            var wbPart = doc.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            return sheet != null;
        }

        public static MemoryStream ToExcel(this DataSet ds)
        {
            var memoryStream = new MemoryStream();
            using (var workbook = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart!.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                // add styles
                workbook.AddStyleSheet();

                uint sheetId = 1;

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData); 

                    var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets!.Elements<Sheet>().Any())
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId!.Value).Max() + 1;

                    var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };

                    sheets.Append(sheet);

                    var headerRow = new Row(); 

                    var columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(column.Caption),
                            StyleIndex = Convert.ToUInt32(column.DefaultValue), 
                        };
                        headerRow.AppendChild(cell);
                    } 

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row();
                        foreach (var col in columns)
                        {
                            var (cellType, value, type) = TypeFinder.Get(dsrow[col]!);
                            var cell = new Cell
                            {
                                DataType = cellType,
                                CellValue = new CellValue((dynamic)value) 
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