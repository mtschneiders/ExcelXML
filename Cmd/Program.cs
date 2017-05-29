using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Diagnostics;
using ExcelXML;

namespace Cmd
{
    class Program
    {
        private static string _basePath;
        
        private static Dictionary<string, int> _sharedStrings = new Dictionary<string, int>();
        private static string _sheetDataHeader;
        private static string _sheetDataFooter;
        private static List<List<object>> _data;// = GetData();

        public static void DoStuff()
        {
            string path = @"C:\Users\Mateus\Desktop\stuff\Book1.xlsx";
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            _basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());
            using (var file = ExcelFile.LoadFromTemplate(path))
            {
                file.BeginWritingData();

                List<List<object>> data = GetData();

                foreach (var rowValues in data)
                    file.WriteRow(rowValues);

                file.EndWritingData();
                file.SaveAs(_basePath + ".xlsx");
            }
        }
        
        public static void Main(string[] args)
        {
            DoStuff();
            return;

            string path = @"C:\Users\Mateus\Desktop\stuff\Book1.xlsx";
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            var tempInfo = new DirectoryInfo(tempBase);
            foreach (var item in tempInfo.GetFileSystemInfos())
            {
                if (item is DirectoryInfo)
                    ((DirectoryInfo)item).Delete(true);
                else
                    item.Delete();
            }

            var stopwatch = Stopwatch.StartNew();
            var stopwatchLoad = Stopwatch.StartNew();
            _basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());
            ZipFile.ExtractToDirectory(path, _basePath);

            LoadSharedStrings();
            LoadSheetSections();
            stopwatchLoad.Stop();

            var stopwatchWriteSheetData = Stopwatch.StartNew();

            string sheet1File = Path.Combine(_basePath, "xl", "worksheets", "sheet1.xml");
            using (var stream = new FileStream(sheet1File, FileMode.Create, FileAccess.ReadWrite))
            using (var writer = new StreamWriter(stream, Encoding.UTF8, 65536))
            {
                int rowCount = _sheetDataHeader.Split(new[] { "</row>" }, StringSplitOptions.None).Length - 1;
                writer.Write(_sheetDataHeader);
                
                foreach (var rowValues in _data)
                {
                    WriteRow(writer, rowValues, rowCount + 1);
                    rowCount++;
                }

                writer.Write(_sheetDataFooter);
            }
            stopwatchWriteSheetData.Stop();

            var stopwatchWriteSharedStrings = Stopwatch.StartNew();
            SaveSharedStrings();
            stopwatchWriteSharedStrings.Stop();
            var stopwatchZip = Stopwatch.StartNew();
            ZipFile.CreateFromDirectory(_basePath, _basePath + ".xlsx");
            stopwatchZip.Stop();
            stopwatch.Stop();
            Console.WriteLine("Load: " + stopwatchLoad.ElapsedMilliseconds);
            Console.WriteLine("WriteData: " + stopwatchWriteSheetData.ElapsedMilliseconds);
            Console.WriteLine("WriteValues: " + stopwatchWriteValues.ElapsedMilliseconds);
            Console.WriteLine("WriteSharedStrings: " + stopwatchWriteSharedStrings.ElapsedMilliseconds);
            Console.WriteLine("Zip: " + stopwatchZip.ElapsedMilliseconds);
            Console.WriteLine("Total: " + stopwatch.ElapsedMilliseconds);
            _sharedStrings = null;
            _sheetDataHeader = null;
            _sheetDataFooter = null;
            _data = null;
            GC.Collect();

            Console.Read();
        }

        private static List<List<object>> GetData()
        {
            List<List<object>> data = new List<List<object>>();

            var random = new Random();
            for (int i = 0; i < 100; i++)
            {
                int lineNumber = i + 3;
                var x = random.Next(0, 10);
                var y = Math.Round(random.NextDouble(), 2, MidpointRounding.AwayFromZero);
                data.Add(new List<object>
                {
                    x,"CD BLABLA BLA BLA BLA "+i, x,"REDE BLA BLA BLA BLA"+i,x,"BRAHMA LATA 350 BLA BLA BLACX12"+i,y,y,y,y,y,y, $"=(I{lineNumber}+J{lineNumber})*(H{lineNumber}/100)"
                });
            }

            return data;
        }
        
        private static Stopwatch stopwatchWriteValues = new Stopwatch();
        
        private static void WriteRow(StreamWriter writer, List<object> values, int rowNumber)
        {
            writer.Write($"<row r=\"{rowNumber}\" spans=\"1:{values.Count}\" x14ac:dyDescent=\"0.25\">");

            for (int i = 0; i < values.Count; i++)
            {
                stopwatchWriteValues.Start();
                object value = values[i];
                int columnNumber = i + 1;
                string columnName = GetExcelColumnName(columnNumber);
                string r = $"{columnName}{rowNumber}";
                writer.Write($"<c r=\"{r}\"");

                if (value is string)
                    writer.Write(" t=\"s\"");

                if (value is DateTime)
                    writer.Write(" s=\"1\"");


                writer.Write("><v>");

                if (value is string)
                    writer.Write(_sharedStrings.GetValueOrNew(value as string));
                else if (value is DateTime)
                    writer.Write((int)((DateTime)value).ToOADate());
                else if (value is double)
                    writer.Write(((double)value).ToString(CultureInfo.InvariantCulture));
                else if (value is decimal)
                    writer.Write(((decimal)value).ToString(CultureInfo.InvariantCulture));
                else
                    writer.Write(value);

                writer.Write("</v></c>");
                stopwatchWriteValues.Stop();
            }
            writer.Write("</row>");
        }
        
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static void LoadSheetSections()
        {
            string sheet1File = Path.Combine(_basePath, "xl", "worksheets", "sheet1.xml");
            var sheetDataText= File.ReadAllText(sheet1File);
            var splitExpression = "</sheetData>";
            var sheetDataValues = sheetDataText.Split(new[] { "</sheetData>" }, StringSplitOptions.None);
            _sheetDataHeader = sheetDataValues[0];
            _sheetDataFooter = splitExpression + sheetDataValues[1];
        }

        private static void SaveSharedStrings()
        {
            string sharedStringsFile = Path.Combine(_basePath, "xl", "sharedStrings.xml");            
            using (var stream = new FileStream(sharedStringsFile, FileMode.Create, FileAccess.Write))
            using(var sw = new StreamWriter(stream, Encoding.UTF8, 65536))
            {
                sw.Write(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count));
                foreach (var item in _sharedStrings)
                {
                    string t = item.Key;
                    if (t.Length > 0 && (t[0] == ' ' || t[t.Length - 1] == ' ' || t.Contains("  ") || t.Contains("\t") || t.Contains("\n") || t.Contains("\n")))
                    {
                        sw.Write("<si><t xml:space=\"preserve\">");
                    }
                    else
                    {
                        sw.Write("<si><t>");
                    }

                    sw.Write(ConvertUtil.ExcelEscapeString(t));
                    sw.Write("</t></si>");
                }
                sw.Write("</sst>");
            }
        }


        private static void LoadSharedStrings()
        {
            string sharedStringsFile = Path.Combine(_basePath, "xl", "sharedStrings.xml");
            var x = XDocument.Load(sharedStringsFile);

            var idx = 0;
            foreach (var item in x.Descendants(XName.Get("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")))
            {
                _sharedStrings.Add(item.Value, idx);
                idx++;
            }
        }
    }


}
