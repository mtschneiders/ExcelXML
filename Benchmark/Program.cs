using System;
using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Attributes.Jobs;
using System.Linq;
using SimpleXL;
using OfficeOpenXml;

namespace Benchmark
{
    [MemoryDiagnoser]
    [ShortRunJob]
    public class ExcelXMLBM
    {
        private List<List<object>> _simpleXLData;
        private List<object[]> _epplusData;
        private string _tempDirectory;
        private const string COSNT_DUMMY_STRING = "IODJSAOIJ@OIDJASOIJONOJBOPAINEPIOQBWNI";

        [Params(10000, 100000)]
        public int NumRecords { get;set; }

        [Params(10)]
        public int NumColumnsString { get; set; }

        [Params(10)]
        public int NumColumnsNumber { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            _simpleXLData = GetData().ToList();
            _epplusData = GetEpplusData().ToList();
            _tempDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");

            if (Directory.Exists(_tempDirectory))
                new DirectoryInfo(_tempDirectory).Delete(true);

            Directory.CreateDirectory(_tempDirectory);
        }

        [Benchmark]
        public void SimpleXL()
        {
            string filePath = Path.Combine(_tempDirectory, Guid.NewGuid().ToString());
            using (var file = new XLFile())
            {
                file.WriteData(_simpleXLData);
                file.SaveAs(filePath + ".xlsx");
            }
        }

        [Benchmark]
        public void EPPlus()
        {
            string filePath = Path.Combine(_tempDirectory, Guid.NewGuid().ToString());
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].LoadFromArrays(_epplusData);
                package.Save();
            }
        }

        public IEnumerable<object[]> GetEpplusData()
        {
            var random = new Random();
            for (int i = 0; i < NumRecords; i++)
            {
                var objectlist = new List<object>();

                for (int j = 0; j < NumColumnsString; j++)
                    objectlist.Add(COSNT_DUMMY_STRING + j);

                for (int j = 0; j < NumColumnsNumber; j++)
                {
                    var x = random.Next(0, 100);
                    objectlist.Add(x);
                }

                yield return objectlist.ToArray();
            }
        }

        public IEnumerable<List<object>> GetData()
        {
            var random = new Random();
            for (int i = 0; i < NumRecords; i++)
            {
                var objectlist = new List<object>();

                for (int j = 0; j < NumColumnsString; j++)
                    objectlist.Add(COSNT_DUMMY_STRING+j);

                for (int j = 0; j < NumColumnsNumber; j++)
                {
                    var x = random.Next(0, 100);
                    objectlist.Add(x);
                }

                yield return objectlist;
            }
        }
    }

    class Program
    {
        [Benchmark]
        public static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<ExcelXMLBM>();
        }
    }
}
