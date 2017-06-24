using SimpleXL;
using System;
using System.Collections.Generic;
using System.Data;
using Xunit;
using System.IO;

namespace Tests
{
    public class XLFileTests
    {
        [Fact]
        public void SaveWithInvalidPath()
        {
            using (XLFile file = new XLFile(new VirtualFileSystem()))
            {
                Assert.Throws<ArgumentNullException>(() => file.SaveAs(string.Empty));
                Assert.Throws<ArgumentNullException>(() => file.SaveAs(null));
                Assert.Throws<ArgumentNullException>(() => file.SaveAs(" "));
                Assert.Throws<ArgumentException>(() => file.SaveAs("randomfolder/randomfile"));
                Assert.Throws<ArgumentException>(() => file.SaveAs("/"));
            }
        }

        [Fact]
        public void CreateAllDirectories()
        {
            string basePath = @"C:\TempXLPath";
            string folderName = "XLFolder";
            var fileSystem = new VirtualFileSystem();
            using (XLFile file = new XLFile(fileSystem, basePath, folderName))
            {
                string folderPath = Path.Combine(basePath, folderName);
                file.SaveAs(@"C:\workbook.xlsx");

                Assert.True(fileSystem.Directories.Contains(folderPath));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "_rels")));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "docProps")));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "xl")));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "xl", "_rels")));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "xl", "theme")));
                Assert.True(fileSystem.Directories.Contains(Path.Combine(folderPath, "xl", "worksheets")));
                Assert.Equal(7, fileSystem.Directories.Count);
            }
        }

        [Fact]
        public void CreateZipFile()
        {
            string basePath = @"C:\TempXLPath";
            string folderName = "XLFolder";
            var fileSystem = new VirtualFileSystem();
            using (XLFile file = new XLFile(fileSystem, basePath, folderName))
            {
                string folderPath = Path.Combine(basePath, folderName);
                string xlFile = @"C:\workbook.xlsx";
                file.SaveAs(xlFile);
                Assert.True(fileSystem.ZipFiles.ContainsKey(folderPath));
                Assert.True(fileSystem.ZipFiles.ContainsValue(xlFile));
            }
        }

        [Fact]
        public void CreateAllFiles()
        {
            string basePath = @"C:\TempXLPath";
            string folderName = "XLFolder";
            var fileSystem = new VirtualFileSystem();
            using (XLFile file = new XLFile(fileSystem, basePath, folderName))
            {
                string folderPath = Path.Combine(basePath, folderName);
                string xlFile = @"C:\workbook.xlsx";
                file.SaveAs(xlFile);
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "xl", "sharedStrings.xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "xl", "styles.xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "xl", "workbook.xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "xl", "_rels", "workbook.xml.rels")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "xl", "theme", "theme1.xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "[Content_Types].xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "_rels", ".rels")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "docProps", "app.xml")));
                Assert.True(fileSystem.Files.Contains(Path.Combine(folderPath, "docProps", "core.xml")));
                Assert.Equal(9, fileSystem.Files.Count);
            }
        }

        [Fact]
        public void DeleteTempDirectory()
        {
            string basePath = @"C:\TempXLPath";
            string folderName = "XLFolder";
            var fileSystem = new VirtualFileSystem();
            using (XLFile file = new XLFile(fileSystem, basePath, folderName))
            {
                string folderPath = Path.Combine(basePath, folderName);
                string xlFile = @"C:\workbook.xlsx";
                file.SaveAs(xlFile);
                Assert.True(fileSystem.DeletedDirectories.Contains(folderPath));
            }
        }

        [Fact]
        public void WriteWithNoData()
        {
            using (XLFile file = new XLFile())
            {
                Assert.Null(Record.Exception(() => file.WriteData(null)));
                Assert.Null(Record.Exception(() => file.WriteData((DataTable)null)));
                Assert.Null(Record.Exception(() => file.WriteData(null, true)));
                Assert.Null(Record.Exception(() => file.WriteData(null, false)));
                Assert.Null(Record.Exception(() => file.WriteData(new List<List<object>>(){ null })));
                Assert.Null(Record.Exception(() => file.WriteData(new List<List<object>>() { new List<object> { null } })));
            }
        }

        [Fact]
        public void ConfigureInvalidRanges()
        {
            using (XLFile file = new XLFile())
            {
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(null, null));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(string.Empty, null));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(" ", null));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(null, new XLRangeConfig()));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(string.Empty, new XLRangeConfig()));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange(" ", new XLRangeConfig()));
                Assert.Throws<ArgumentNullException>(() => file.ConfigureRange("A1:B1", null));

                Assert.Throws<ArgumentException>(() => file.ConfigureRange(":", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A:", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange(":A", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A1:", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A1:B", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A:B1", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("AAAA1:B1", new XLRangeConfig()));
                Assert.Throws<ArgumentException>(() => file.ConfigureRange("A1:BAAAA1", new XLRangeConfig()));
            }
        }


        [Fact]
        public void ConfigureValidRanges()
        {
            using (XLFile file = new XLFile())
            {
                Assert.Null(Record.Exception(() => file.ConfigureRange("A1:B1", new XLRangeConfig())));
                Assert.Null(Record.Exception(() => file.ConfigureRange("A11:B11", new XLRangeConfig())));
                Assert.Null(Record.Exception(() => file.ConfigureRange("A1:B11111", new XLRangeConfig())));
            }
        }
    }
}
