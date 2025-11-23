using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Mahamudra.Excel.Infrastructure
{
    /// <summary>
    /// Builds Excel stylesheet with fonts, fills, borders, and cell formats.
    /// </summary>
    internal static class StylesheetBuilder
    {
        internal static WorkbookStylesPart AddStyleSheet(SpreadsheetDocument spreadsheet)
        {
            var stylesheet = spreadsheet.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
            var workbookstylesheet = new Stylesheet();

            uint DATETIME_FORMAT = 164;
            uint DIGITS4_FORMAT = 165;

            var numberingFormats = new NumberingFormats();
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(DATETIME_FORMAT),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            });
            numberingFormats.Append(new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(DIGITS4_FORMAT),
                FormatCode = StringValue.FromString("0000")
            });
            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);

            var fonts = CreateFonts();
            var fills = CreateFills();
            var borders = CreateBorders();
            var cellStyleFormats = CreateCellStyleFormats();
            var cellFormats = CreateCellFormats(DATETIME_FORMAT, DIGITS4_FORMAT);

            workbookstylesheet.Append(fonts);
            workbookstylesheet.Append(fills);
            workbookstylesheet.Append(borders);
            workbookstylesheet.Append(cellFormats);

            stylesheet.Stylesheet = workbookstylesheet;
            stylesheet.Stylesheet.Save();

            return stylesheet;
        }

        private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new ForegroundColor
            {
                Rgb = new HexBinaryValue
                {
                    Value = string.Format("{0:X2}{1:X2}{2:X2}{3:X3}", fillColor.R, fillColor.G, fillColor.B, fillColor.A)
                }
            };
        }

        private static Fonts CreateFonts()
        {
            var fonts = new Fonts();
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) }
            });
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font
            {
                FontName = new FontName { Val = StringValue.FromString("Arial") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) },
                Bold = new Bold()
            });
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            return fonts;
        }

        private static Fills CreateFills()
        {
            var fills = new Fills();
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            fills.Append(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });
            fills.Append(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Append(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = TranslateForeground(System.Drawing.Color.LightSkyBlue),
                    BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb }
                }
            });
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            return fills;
        }

        private static Borders CreateBorders()
        {
            var borders = new Borders();
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            return borders;
        }

        private static CellStyleFormats CreateCellStyleFormats()
        {
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            });
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            return cellStyleFormats;
        }

        private static CellFormats CreateCellFormats(uint datetimeFormat, uint digits4Format)
        {
            var cellFormats = new CellFormats();

            // Index 0: Default
            cellFormats.Append(new CellFormat { FontId = 0, FillId = 0, BorderId = 0, NumberFormatId = 0, FormatId = 0 });

            // Index 1: Wrapper
            cellFormats.Append(new CellFormat { Alignment = new Alignment { WrapText = true } });

            // Index 2: Standard Date
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 14,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 3: Standard Number Decimal
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 4: Standard DateTime
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = datetimeFormat,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 5: Standard Integer
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 3,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 6: Percent
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 10,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 7: Digit4
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = digits4Format,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // Index 8: Header
            cellFormats.Append(new CellFormat
            {
                NumberFormatId = 49,
                FontId = 1,
                FillId = 0,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                Alignment = new Alignment { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }
            });

            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            return cellFormats;
        }
    }
}
