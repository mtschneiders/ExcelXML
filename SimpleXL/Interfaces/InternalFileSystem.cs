using System.IO;
using System.IO.Compression;
using System.Text;

namespace SimpleXL.Interfaces
{
    internal class InternalFileSystem : IFileSystem
    {
        private const int WRITE_BUFFER_SIZE = 65536;
        
        public TextWriter CreateNewFile(string filePath)
        {
            var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            return new StreamWriter(fileStream, Encoding.UTF8, WRITE_BUFFER_SIZE);
        }

        public void WriteAllText(string path, string contents) => File.WriteAllText(path, contents);
        public void CreateZipFromDirectory(string sourceDirectoryName, string destinationArchiveFileName) => ZipFile.CreateFromDirectory(sourceDirectoryName, destinationArchiveFileName);
        public void CreateDirectory(string path) => Directory.CreateDirectory(path);
        public bool DirectoryExists(string path) => Directory.Exists(path);
    }
}
