using System.IO;

namespace SimpleXL.Extensions
{
    internal static class StringExtensions
    {
        public static bool IsValidFilePath(this string value)
        {
            try { new FileInfo(value); }
            catch { return false; }
            return true;
        }

        public static FileInfo GetFileInfo(this string value)
        {
            try { return new FileInfo(value); }
            catch { }
            return null;
        }
    }
}
