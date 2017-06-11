using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace SimpleXL
{
    public class ExcelFile : IDisposable
    {
        private const int DEFAULT_WRITE_BUFFER_SIZE = 65536;
        private const string XL_FOLDER = "xl";
        private const string WORKSHEETS_FOLDER = "worksheets";

        private string _templatePath;
        private string _basePath;
        private Dictionary<string, int> _sharedStrings = new Dictionary<string, int>();
        private string _sheetDataHeader;
        private string _sheetDataFooter;
        private FileStream _fileStream;
        private StreamWriter _streamWriter;
        private int _currentRowCount;

        private ExcelFile(string templatePath)
        {
            _templatePath = templatePath;
            _basePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        }

        public static ExcelFile LoadFromTemplate(string templatePath)
        {
            ExcelFile file = new ExcelFile(templatePath);
            file.Load();
            return file;
        }

        private void Load()
        {
            ZipHelper.ExtractToDirectory(_templatePath, _basePath);
            LoadSharedStrings();
            LoadSheetSections();
        }

        public void BeginWritingData()
        {
            string sheet1File = Path.Combine(_basePath, XL_FOLDER, WORKSHEETS_FOLDER, "sheet1.xml");
            _fileStream = new FileStream(sheet1File, FileMode.Create, FileAccess.ReadWrite);
            _streamWriter = new StreamWriter(_fileStream, Encoding.UTF8, DEFAULT_WRITE_BUFFER_SIZE);
            _streamWriter.Write(_sheetDataHeader);
        }
        
        public void WriteRow(List<object> values)
        {
            _currentRowCount++;
            _streamWriter.Write("<row r=\"");
            _streamWriter.Write(_currentRowCount);
            _streamWriter.Write("\">");
            
            for (int i = 0; i < values.Count; i++)
            {
                object value = values[i];
                string columnName = ExcelHelper.GetExcelColumnName(i + 1);
                
                _streamWriter.Write("<c r=\"");
                _streamWriter.Write(columnName);
                _streamWriter.Write(_currentRowCount);
                _streamWriter.Write("\"");

                if (value is string)
                    _streamWriter.Write(" t=\"s\"");

                if (value is DateTime)
                    _streamWriter.Write(" s=\"1\"");

                _streamWriter.Write(">");
                
                if (value is string valueStr)
                    WriteStringValueTag(_streamWriter, valueStr);
                else if (value is DateTime valueDateTime)
                    WriteValueTag(_streamWriter, valueDateTime.ToOADate());
                else if (value is double valueDouble)
                    WriteValueTag(_streamWriter, valueDouble.ToString(CultureInfo.InvariantCulture));
                else if (value is decimal valueDecimal)
                    WriteValueTag(_streamWriter, valueDecimal.ToString(CultureInfo.InvariantCulture));
                else
                    WriteValueTag(_streamWriter, value);
                    
                _streamWriter.Write("</c>");
            }

            _streamWriter.Write("</row>");
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
                
        public void EndWritingData()
        {
            _streamWriter.Write(_sheetDataFooter);
        }

        public void SaveAs(string filePath)
        {
            _streamWriter?.Dispose();
            _fileStream?.Dispose();
            SaveSharedStrings();
            ZipFile.CreateFromDirectory(_basePath, filePath);
            new DirectoryInfo(_basePath).Delete(true);
        }

        private void LoadSharedStrings()
        {
            string sharedStringsFile = Path.Combine(_basePath, XL_FOLDER, "sharedStrings.xml");
            var document = XDocument.Load(sharedStringsFile);
            foreach (var item in document.Descendants(XName.Get("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")))
                _sharedStrings.Add(item.Value, _sharedStrings.Keys.Count);
        }

        private void SaveSharedStrings()
        {
            string sharedStringsFile = Path.Combine(_basePath, XL_FOLDER, "sharedStrings.xml");
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

        private void LoadSheetSections()
        {
            string sheet1File = Path.Combine(_basePath, XL_FOLDER, WORKSHEETS_FOLDER, "sheet1.xml");
            var sheetDataText = File.ReadAllText(sheet1File);
            var splitExpression = "</sheetData>";
            var sheetDataValues = sheetDataText.Split(new[] { splitExpression }, StringSplitOptions.None);
            _sheetDataHeader = sheetDataValues[0];
            _sheetDataFooter = splitExpression + sheetDataValues[1];
            _currentRowCount = _sheetDataHeader.Split(new[] { "</row>" }, StringSplitOptions.None).Length - 1;
        }
        
        public void Dispose()
        {
            _streamWriter?.Dispose();
            _fileStream?.Dispose();
            _templatePath = null;
            _basePath = null;
            _sharedStrings = null;
            _sheetDataHeader = null;
            _sheetDataFooter = null;
        }

    }
}
