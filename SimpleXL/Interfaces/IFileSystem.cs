using System;
using System.IO;

namespace SimpleXL.Interfaces
{
    internal interface IFileSystem
    {
        void CreateDirectory(string path);
        bool DirectoryExists(string path);
        void DeleteDirectory(string path);
        TextWriter CreateNewFile(string filePath);
        void WriteAllText(string path, string contents);
        void CreateZipFromDirectory(string sourceDirectoryName, string destinationArchiveFileName);
    }
}
