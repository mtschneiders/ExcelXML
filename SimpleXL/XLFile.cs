using ExcelXML.Extensions;
using SimpleXL.Extensions;
using SimpleXL.Helpers;
using SimpleXL.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SimpleXL
{
    /// <summary> Class responsible for building and saving Excel Files
    /// </summary>
    public class XLFile : IDisposable
    {
        private string _temporaryBasePath;
        private Dictionary<string, int> _sharedStrings = new Dictionary<string, int>();
        private Dictionary<XLRange, XLRangeConfig> _rangeConfigs = new Dictionary<XLRange, XLRangeConfig>();
        private List<XLStyle> _styles = new List<XLStyle>();
        private IFileSystem _fileSystem;

        /// <summary> Creates a new XLFile instance
        /// </summary>
        public XLFile() : this(new InternalFileSystem()) { }

        internal XLFile(IFileSystem fileSystem) : this(fileSystem, GetBasePath(), Guid.NewGuid().ToString()) { }
        internal XLFile(IFileSystem fileSystem, string basePath, string folderName)
        {
            _temporaryBasePath = Path.Combine(basePath, folderName);
            _fileSystem = fileSystem;
            _styles.Add(new XLStyle());
        }

        private static string GetBasePath()
        {
#if DEBUG
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
#else
            return Path.GetTempPath();
#endif
        }

        private string GetXLFilePath(string fileName) => Path.Combine(_temporaryBasePath, "xl", fileName);
        private string GetWorksheetsFilePath(string fileName) => Path.Combine(_temporaryBasePath, "xl", "worksheets", fileName);
        
        private void EnsureFolderStructureIsCreated()
        {
            if (!_fileSystem.DirectoryExists(_temporaryBasePath))
            {
                _fileSystem.CreateDirectory(_temporaryBasePath);
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "_rels"));
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "docProps"));
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "xl"));
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "xl", "_rels"));
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "xl", "theme"));
                _fileSystem.CreateDirectory(Path.Combine(_temporaryBasePath, "xl", "worksheets"));
            }
        }

        /// <summary> Writes data to the excel package
        /// </summary>
        /// <param name="table">DataTable</param>
        /// <param name="writeColumnHeaders">Indicates if the column headers should be written to the file</param>
        public void WriteData(DataTable table, bool writeColumnHeaders = true)
        {
            if (table == null)
                return;

            WriteSheetData((writer) =>
            {
                int rowNumber = 1;
                if (writeColumnHeaders)
                {
                    var columnNames = table.Columns.OfType<DataColumn>().Select(col => col.Caption ?? col.ColumnName).ToArray();
                    WriteRow(writer, columnNames, rowNumber);
                    rowNumber++;
                }

                foreach (DataRow row in table.Rows)
                {
                    WriteRow(writer, row.ItemArray, rowNumber);
                    rowNumber++;
                }
            });
        }

        /// <summary> Writes data to the excel package
        /// </summary>
        /// <param name="data">Collection of object collections</param>
        public void WriteData(IEnumerable<IEnumerable<object>> data)
        {
            if (data == null)
                return;

            WriteSheetData((writer) =>
            {
                int rowNumber = 1;
                foreach (var values in data)
                {
                    WriteRow(writer, values?.ToArray(), rowNumber);
                    rowNumber++;
                }
            });
        }

        private void WriteSheetData(Action<TextWriter> dataWriteAction)
        {
            EnsureFolderStructureIsCreated();
            string sheet1File = GetWorksheetsFilePath("sheet1.xml");
            using (var writer = _fileSystem.CreateNewFile(sheet1File))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
                writer.Write("<dimension ref=\"A1\"/>");
                writer.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>");
                writer.Write("<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>");
                writer.Write("<sheetData>");
                dataWriteAction(writer);
                writer.Write("</sheetData>");
                writer.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
                writer.Write("</worksheet>");
            }
        }
        
        private void WriteRow(TextWriter writer, object[] values, int rowNumber)
        {
            writer.Write("<row r=\"");
            writer.Write(rowNumber);
            writer.Write("\">");

            if (values != null)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    object value = values[i];
                    int columnNumber = i + 1;
                    string columnName = ExcelHelper.GetExcelColumnName(columnNumber);

                    writer.Write("<c r=\"");
                    writer.Write(columnName);
                    writer.Write(rowNumber);
                    writer.Write("\"");

                    if (value is string)
                        writer.Write(" t=\"s\"");

                    int? styleId = GetStyleId(columnNumber, rowNumber);
                    if (styleId.HasValue)
                    {
                        writer.Write(" s=\"");
                        writer.Write(styleId.Value);
                        writer.Write("\"");
                    }

                    writer.Write(">");

                    if (value is string valueStr)
                        WriteStringValueTag(writer, valueStr);
                    else if (value is DateTime valueDateTime)
                        WriteValueTag(writer, valueDateTime.ToOADate());
                    else if (value is double valueDouble)
                        WriteValueTag(writer, valueDouble.ToString(CultureInfo.InvariantCulture));
                    else if (value is decimal valueDecimal)
                        WriteValueTag(writer, valueDecimal.ToString(CultureInfo.InvariantCulture));
                    else
                        WriteValueTag(writer, value);

                    writer.Write("</c>");
                }
            }

            writer.Write("</row>");
        }

        private void WriteValueTag(TextWriter writer, object value)
        {
            writer.Write("<v>");
            writer.Write(value);
            writer.Write("</v>");
        }

        private void WriteStringValueTag(TextWriter writer, string value)
        {
            if (value.StartsWith("="))
            {
                writer.Write("<f>");
                writer.Write(value);
                writer.Write("</f>");
            }
            else
                WriteValueTag(writer, _sharedStrings.GetValueOrNew(value));
        }
        
        private int? GetStyleId(int columnNumber, int rowNumber)
        {
            foreach (var item in _rangeConfigs)
            {
                var range = item.Key;

                if(columnNumber.Between(range.ColumnNumberStart, range.ColumnNumberEnd) && rowNumber.Between(range.RowNumberStart, range.RowNumberEnd))
                    return range.StyleId;
            }

            return null;
        }

        /// <summary> Defines a configuration for a range of cells
        /// </summary>
        /// <param name="range"></param>
        /// <param name="config"></param>
        public void ConfigureRange(string range, XLRangeConfig config)
        {
            if (string.IsNullOrWhiteSpace(range))
                throw new ArgumentNullException(nameof(range));

            if (config == null)
                throw new ArgumentNullException(nameof(config));
            
            int styleId = AddNewRangeStyle(config);
            _rangeConfigs[new XLRange(range, styleId)] = config;
        }

        private int AddNewRangeStyle(XLRangeConfig config)
        {
            for (int i = 0; i < _styles.Count; i++)
            {
                var style = _styles[i];
                if (style.BorderId == config.Border.ToInt() &&
                   style.FontId == (int)config.Font &&
                   style.NumFormatId == (int)config.Format)
                    return i;
            }
            
            _styles.Add(new XLStyle()
            {
                BorderId = config.Border.ToInt(),
                FontId = (int)config.Font,
                NumFormatId = (int)config.Format
            });
            return _styles.Count - 1;
        }

        /// <summary> Saves the excel file to the specified path
        /// </summary>
        /// <param name="filePath">Absolute file path</param>
        public void SaveAs(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            FileInfo fileInfo = filePath.GetFileInfo();
            
            if (fileInfo == null || fileInfo.Directory == null || !fileInfo.Directory.Exists)
                throw new ArgumentException("Invalid path", nameof(filePath));

            EnsureFolderStructureIsCreated();
            SaveSharedStrings();
            SaveStyles();
            SaveEmbeddedFiles();
            _fileSystem.CreateZipFromDirectory(_temporaryBasePath, filePath);
            
#if RELEASE
            new DirectoryInfo(_temporaryBasePath).Delete(true);
#endif 
        }
        
        private void SaveEmbeddedFiles()
        {
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "[Content_Types].xml"), Resources.Template_Content_Types_);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "_rels", ".rels"), Resources.Template_rels_rels);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "docProps", "app.xml"), Resources.Template_docProps_app);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "docProps", "core.xml"), Resources.Template_docProps_core);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "xl", "workbook.xml"), Resources.Template_xl_workbook);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "xl", "_rels", "workbook.xml.rels"), Resources.Template_xl_rels_workbook);
            _fileSystem.WriteAllText(Path.Combine(_temporaryBasePath, "xl", "theme", "theme1.xml"), Resources.Template_xl_theme_theme1);
        }

        private void SaveSharedStrings()
        {
            string sharedStringsFile = GetXLFilePath("sharedStrings.xml");
            using (var writer = _fileSystem.CreateNewFile(sharedStringsFile))
            {
                writer.Write(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count));
                foreach (var item in _sharedStrings)
                {
                    string t = item.Key;
                    if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\t") || t.Contains("\n") || t.Contains("\n")))
                        writer.Write("<si><t xml:space=\"preserve\">");
                    else
                        writer.Write("<si><t>");

                    writer.Write(ExcelHelper.ExcelEscapeString(t));
                    writer.Write("</t></si>");
                }
                writer.Write("</sst>");
            }
        }

        private void SaveStyles()
        {
            string stylesFile = GetXLFilePath("styles.xml");
            using (var writer = _fileSystem.CreateNewFile(stylesFile))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\">");
                writer.Write(Resources.styles_fonts);
                writer.Write("<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>");
                writer.Write(Resources.styles_borders);
                writer.Write("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
                writer.Write($"<cellXfs count=\"{_styles.Count}\">");

                for (int i = 0; i < _styles.Count; i++)
                {
                    var style = _styles[i];
                    writer.Write("<xf numFmtId=\"");
                    writer.Write(style.NumFormatId);
                    writer.Write("\" fillId=\"0\" borderId=\"");
                    writer.Write(style.BorderId);
                    writer.Write("\" xfId=\"0\" fontId=\"");
                    writer.Write(style.FontId);
                    writer.Write("\" ");
                    if (style.NumFormatId > 0)
                        writer.Write("applyNumberFormat=\"1\" ");

                    if (style.FontId > 0)
                        writer.Write("applyFont=\"1\" ");

                    if (style.BorderId > 0)
                        writer.Write("applyBorder=\"1\"");

                    writer.Write("/>");
                }

                writer.Write("</cellXfs>");
                writer.Write("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/>");
                writer.Write("</styleSheet>");
            }
        }
        
        /// <summary> Disposes of an XLFile instance
        /// </summary>
        public void Dispose()
        {
#if RELEASE
            if (Directory.Exists(_temporaryBasePath))
                new DirectoryInfo(_temporaryBasePath).Delete(true);
#endif 

            _temporaryBasePath = null;
            _sharedStrings = null;
            _styles = null;
            _rangeConfigs = null;
        }
    }
}
