using System.Data;
using System.IO;

namespace Mahamudra.Excel.Abstractions
{
    /// <summary>
    /// Interface for writing data to Excel files.
    /// </summary>
    public interface IExcelWriter
    {
        /// <summary>
        /// Converts a DataSet to an Excel file stream.
        /// </summary>
        /// <param name="dataSet">The DataSet containing data to export.</param>
        /// <returns>A MemoryStream containing the Excel file.</returns>
        MemoryStream Write(DataSet dataSet);
    }
}
