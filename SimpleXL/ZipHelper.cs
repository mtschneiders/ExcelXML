using System;
using System.IO;
using System.IO.Compression;

namespace ExcelXML
{
    public static class ZipHelper
    {
        /// <summary> Extract a zip file using a customized buffersize when reading the compressed files into memory
        /// </summary>
        /// <param name="sourceArchiveFileName"></param>
        /// <param name="destinationDirectoryName"></param>
        /// <param name="bufferSize"></param>
        public static void ExtractToDirectory(string sourceArchiveFileName, string destinationDirectoryName, int bufferSize = 1024)
        {
            using (var fileStream = File.Open(sourceArchiveFileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var source = new ZipArchive(fileStream, ZipArchiveMode.Read, false, null))
                {
                    DirectoryInfo directoryInfo = Directory.CreateDirectory(destinationDirectoryName);
                    string fullName = directoryInfo.FullName;

                    foreach (ZipArchiveEntry current in source.Entries)
                    {
                        string fullPath = Path.GetFullPath(Path.Combine(fullName, current.FullName));

                        if (!fullPath.StartsWith(fullName, StringComparison.OrdinalIgnoreCase))
                        {
                            throw new IOException("Extracting Zip entry would have resulted in a file outside the specified destination directory");
                        }

                        if (Path.GetFileName(fullPath).Length == 0)
                        {
                            if (current.Length != 0L)
                            {
                                throw new IOException("Zip entry name ends in directory separator character but contains data");
                            }
                            Directory.CreateDirectory(fullPath);
                        }
                        else
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
                            using (Stream stream = File.Open(fullPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                            using (Stream stream2 = current.Open())
                                stream2.CopyTo(stream, bufferSize);
                        }
                    }
                }
            }
        }

    }
}
