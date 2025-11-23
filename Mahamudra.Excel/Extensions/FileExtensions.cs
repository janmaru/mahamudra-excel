using System;
using System.IO;

namespace Mahamudra.Excel.Extensions
{
    /// <summary>
    /// Extension methods for file operations.
    /// </summary>
    public static class FileExtensions
    {
        /// <summary>
        /// Writes a memory stream to a file.
        /// </summary>
        /// <param name="stream">The stream to write.</param>
        /// <param name="filePath">The file path to write to.</param>
        /// <exception cref="ArgumentNullException">Thrown when stream or filePath is null.</exception>
        /// <exception cref="ArgumentException">Thrown when filePath is empty or whitespace.</exception>
        public static void Write(this MemoryStream stream, string filePath)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be null or whitespace.", nameof(filePath));

            using var streamFile = stream;
            var array = streamFile.ToArray();
            File.WriteAllBytes(filePath, array);
        }

        /// <summary>
        /// Checks if a file exists at the specified path.
        /// </summary>
        /// <param name="filePath">The file path to check.</param>
        /// <returns>True if the file exists; otherwise, false.</returns>
        public static bool Exists(this string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// Reads a file into a memory stream.
        /// </summary>
        /// <param name="filePath">The file path to read.</param>
        /// <returns>A MemoryStream containing the file contents.</returns>
        /// <exception cref="ArgumentException">Thrown when filePath is null or whitespace.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static MemoryStream Read(this string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path cannot be null or whitespace.", nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found.", filePath);

            using var fileStream = File.OpenRead(filePath);
            var memStream = new MemoryStream();
            memStream.SetLength(fileStream.Length);
            fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
            return memStream;
        }
    }
}
