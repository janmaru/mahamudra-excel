using System.Data;
using System.IO;

namespace Mahamudra.Excel.Abstractions
{
    /// <summary>
    /// Interface for reading data from Excel files.
    /// </summary>
    public interface IExcelReader
    {
        /// <summary>
        /// Reads an Excel file stream into a DataTable.
        /// </summary>
        /// <param name="stream">The memory stream containing the Excel file.</param>
        /// <returns>A DataTable with the Excel data.</returns>
        DataTable Read(MemoryStream stream);
    }
}
