using System.Collections.Generic;
using System.IO;

namespace Tests
{
    internal class VirtualFileSystem : SimpleXL.Interfaces.IFileSystem
    {
        public Dictionary<string, string> ZipFiles { get; private set; }
        public HashSet<string> Files { get; private set; }
        public HashSet<string> Directories { get; private set; }

        public VirtualFileSystem()
        {
            Files = new HashSet<string>();
            ZipFiles = new Dictionary<string, string>();
            Directories = new HashSet<string>();
        }

        public TextWriter CreateNewFile(string filePath)
        {
            Files.Add(filePath);
            return new StreamWriter(new MemoryStream());
        }

        public void CreateZipFromDirectory(string sourceDirectoryName, string destinationArchiveFileName) => ZipFiles.Add(sourceDirectoryName, destinationArchiveFileName);
        public void WriteAllText(string path, string contents) => Files.Add(path);

        public void CreateDirectory(string path) => Directories.Add(path);
        public bool DirectoryExists(string path) => Directories.Contains(path);
    }
}
