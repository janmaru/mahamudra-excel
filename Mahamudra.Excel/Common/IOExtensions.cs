using System.Data;
using System.IO;
using System.Runtime.CompilerServices;

namespace Mahamudra.Excel.Common
{
    public static class IOExtensions
    {
        public static void  Write(this MemoryStream stream, string filePath)
        {
            using var streamFile = stream;
            var array = streamFile.ToArray();
            File.WriteAllBytes(filePath, array);
        }
        public static bool Exists(this string filePath)
        { 
            return File.Exists(filePath);
        }
        public static MemoryStream Read(this string filePath)
        {
            using var fileStream = File.OpenRead(filePath);
            var memStream = new MemoryStream();
            memStream.SetLength(fileStream.Length);
            fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
            return memStream;
        } 
    }
} 