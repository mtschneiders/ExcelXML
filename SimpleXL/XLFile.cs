using ExcelXML.Extensions;
using SimpleXL.Extensions;
using SimpleXL.Helpers;
using SimpleXL.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace SimpleXL
{
    public class XLFile : IDisposable
    {
        private const int DEFAULT_WRITE_BUFFER_SIZE = 65536;
        
        private string _basePath;
        private Dictionary<string, int> _sharedStrings = new Dictionary<string, int>();
        private Dictionary<XLRange, XLRangeConfig> _rangeConfigs = new Dictionary<XLRange, XLRangeConfig>();
        private List<XLStyle> _styles = new List<XLStyle>();
        private int _currentRowCount;

        public XLFile()
        {
            _basePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            //_basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp", Guid.NewGuid().ToString());
            _styles.Add(new XLStyle());
        }

        private string GetXLFolderPath() => Path.Combine(_basePath, "xl");
        private string GetXLFilePath(string fileName) => Path.Combine(_basePath, "xl", fileName);
        private string GetWorksheetsFilePath(string fileName) => Path.Combine(_basePath, "xl", "worksheets", fileName);
        
        private void EnsureFolderStructureIsCreated()
        {
            if (!Directory.Exists(_basePath))
            {
                Directory.CreateDirectory(_basePath);
                Directory.CreateDirectory(Path.Combine(_basePath, "_rels"));
                Directory.CreateDirectory(Path.Combine(_basePath, "docProps"));
                Directory.CreateDirectory(Path.Combine(_basePath, "xl"));
                Directory.CreateDirectory(Path.Combine(_basePath, "xl", "_rels"));
                Directory.CreateDirectory(Path.Combine(_basePath, "xl", "theme"));
                Directory.CreateDirectory(Path.Combine(_basePath, "xl", "worksheets"));
            }
        }

        public void WriteData(IEnumerable<List<object>> data)
        {
            EnsureFolderStructureIsCreated();

            string sheet1File = GetWorksheetsFilePath("sheet1.xml");
            using (var fileStream = new FileStream(sheet1File, FileMode.Create, FileAccess.ReadWrite))
            using (var streamWriter = new StreamWriter(fileStream, Encoding.UTF8, DEFAULT_WRITE_BUFFER_SIZE))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
                streamWriter.Write("<dimension ref=\"A1\"/>");
                streamWriter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>");
                streamWriter.Write("<sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>");
                streamWriter.Write("<sheetData>");

                foreach (var values in data)
                    WriteRow(streamWriter, values);

                streamWriter.Write("</sheetData>");
                streamWriter.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
                streamWriter.Write("</worksheet>");
            }
        }
        
        private void WriteRow(StreamWriter streamWriter, List<object> values)
        {
            _currentRowCount++;
            streamWriter.Write("<row r=\"");
            streamWriter.Write(_currentRowCount);
            streamWriter.Write("\">");
            
            for (int i = 0; i < values.Count; i++)
            {
                object value = values[i];
                int columnNumber = i + 1;
                string columnName = ExcelHelper.GetExcelColumnName(columnNumber);

                streamWriter.Write("<c r=\"");
                streamWriter.Write(columnName);
                streamWriter.Write(_currentRowCount);
                streamWriter.Write("\"");

                if (value is string)
                    streamWriter.Write(" t=\"s\"");

                int? styleId = GetStyleId(columnNumber, _currentRowCount);
                if (styleId.HasValue)
                {
                    streamWriter.Write(" s=\"");
                    streamWriter.Write(styleId.Value);
                    streamWriter.Write("\"");
                }

                streamWriter.Write(">");
                
                if (value is string valueStr)
                    WriteStringValueTag(streamWriter, valueStr);
                else if (value is DateTime valueDateTime)
                    WriteValueTag(streamWriter, valueDateTime.ToOADate());
                else if (value is double valueDouble)
                    WriteValueTag(streamWriter, valueDouble.ToString(CultureInfo.InvariantCulture));
                else if (value is decimal valueDecimal)
                    WriteValueTag(streamWriter, valueDecimal.ToString(CultureInfo.InvariantCulture));
                else
                    WriteValueTag(streamWriter, value);

                streamWriter.Write("</c>");
            }

            streamWriter.Write("</row>");
        }

        private void WriteValueTag(StreamWriter writer, object value)
        {
            writer.Write("<v>");
            writer.Write(value);
            writer.Write("</v>");
        }

        private void WriteStringValueTag(StreamWriter writer, string value)
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
        
        public int? GetStyleId(int columnNumber, int rowNumber)
        {
            foreach (var item in _rangeConfigs)
            {
                var range = item.Key;

                if(columnNumber.Between(range.ColumnNumberStart, range.ColumnNumberEnd) && rowNumber.Between(range.RowNumberStart, range.RowNumberEnd))
                    return range.StyleId;
            }

            return null;
        }

        public void ConfigureRange(string range, XLRangeConfig config)
        {
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

        public void SaveAs(string filePath)
        {
            SaveSharedStrings();
            SaveStyles();
            SaveCachedFiles();
            ZipFile.CreateFromDirectory(_basePath, filePath);
            new DirectoryInfo(_basePath).Delete(true);
        }
        
        private void SaveCachedFiles()
        {
            File.WriteAllText(Path.Combine(_basePath, "[Content_Types].xml"), Resources.Template_Content_Types_);
            File.WriteAllText(Path.Combine(_basePath, "_rels", ".rels"), Resources.Template_rels_rels);
            File.WriteAllText(Path.Combine(_basePath, "docProps", "app.xml"), Resources.Template_docProps_app);
            File.WriteAllText(Path.Combine(_basePath, "docProps", "core.xml"), Resources.Template_docProps_core);
            File.WriteAllText(Path.Combine(_basePath, "xl", "workbook.xml"), Resources.Template_xl_workbook);
            File.WriteAllText(Path.Combine(_basePath, "xl", "_rels", "workbook.xml.rels"), Resources.Template_xl_rels_workbook);
            File.WriteAllText(Path.Combine(_basePath, "xl", "theme", "theme1.xml"), Resources.Template_xl_theme_theme1);
        }

        private void SaveSharedStrings()
        {
            string sharedStringsFile = GetXLFilePath("sharedStrings.xml");
            using (var stream = new FileStream(sharedStringsFile, FileMode.Create, FileAccess.Write))
            using (var sw = new StreamWriter(stream, Encoding.UTF8, DEFAULT_WRITE_BUFFER_SIZE))
            {
                sw.Write(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count));
                foreach (var item in _sharedStrings)
                {
                    string t = item.Key;
                    if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\t") || t.Contains("\n") || t.Contains("\n")))
                        sw.Write("<si><t xml:space=\"preserve\">");
                    else
                        sw.Write("<si><t>");

                    sw.Write(ExcelHelper.ExcelEscapeString(t));
                    sw.Write("</t></si>");
                }
                sw.Write("</sst>");
            }
        }

        private void SaveStyles()
        {
            string stylesFile = GetXLFilePath("styles.xml");
            using (var stream = new FileStream(stylesFile, FileMode.Create, FileAccess.Write))
            using (var sw = new StreamWriter(stream, Encoding.UTF8, DEFAULT_WRITE_BUFFER_SIZE))
            {
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac x16r2\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\">");
                sw.Write(Resources.styles_fonts);
                sw.Write("<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>");
                sw.Write(Resources.styles_borders);
                sw.Write("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
                sw.Write($"<cellXfs count=\"{_styles.Count}\">");

                for (int i = 0; i < _styles.Count; i++)
                {
                    var style = _styles[i];
                    sw.Write("<xf numFmtId=\"");
                    sw.Write(style.NumFormatId);
                    sw.Write("\" fillId=\"0\" borderId=\"");
                    sw.Write(style.BorderId);
                    sw.Write("\" xfId=\"0\" fontId=\"");
                    sw.Write(style.FontId);
                    sw.Write("\" ");
                    if (style.NumFormatId > 0)
                        sw.Write("applyNumberFormat=\"1\" ");

                    if (style.FontId > 0)
                        sw.Write("applyFont=\"1\" ");

                    if (style.BorderId > 0)
                        sw.Write("applyBorder=\"1\"");

                    sw.Write("/>");
                }

                sw.Write("</cellXfs>");
                sw.Write("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"0\"/>");
                sw.Write("</styleSheet>");
            }
        }
        
        public void Dispose()
        {
            _basePath = null;
            _sharedStrings = null;
            _styles = null;
            _rangeConfigs = null;
        }

    }
}
