namespace Mahamudra.Excel.Domain
{
    /// <summary>
    /// Defines the cell formatting styles available for Excel cells.
    /// </summary>
    public enum XCellStyle
    {
        /// <summary>No specific formatting.</summary>
        None = 0,
        /// <summary>Text wrapping enabled.</summary>
        Wrapper = 1,
        /// <summary>Standard date format (mm-dd-yy).</summary>
        StandardDate = 2,
        /// <summary>Standard decimal number format (#,##0.00).</summary>
        StandardNumberDecimal = 3,
        /// <summary>Standard datetime format (dd/mm/yyyy hh:mm:ss).</summary>
        StandardDateTime = 4,
        /// <summary>Standard integer format (#,##0).</summary>
        StandardInteger = 5,
        /// <summary>Percentage format (0.00%).</summary>
        Percent = 6,
        /// <summary>Four-digit format with leading zeros (0000).</summary>
        Digit4 = 7,
        /// <summary>Header cell formatting with bold text and borders.</summary>
        Header = 8,
    }
}
