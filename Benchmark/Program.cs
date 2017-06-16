using System;
using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Attributes.Jobs;
using System.Linq;

namespace SimpleXL.Benchmark
{
    [MemoryDiagnoser]
    [ShortRunJob]
    public class ExcelXMLBM
    {
        private List<List<object>> _data;
        private const string COSNT_DUMMY_STRING = "IODJSAOIJ@OIDJASOIJONOJBOPAINEPIOQBWNI";

        [Params(10000)]
        public int NumRecords { get;set; }

        [Params(10)]
        public int NumColumnsString { get; set; }

        [Params(10)]
        public int NumColumnsNumber { get; set; }

        [GlobalSetup]
        public void GlobalSetup()
        {
            _data = GetData().ToList();
        }

        [Benchmark]
        public void ExportData()
        {
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            string basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());
            using (var file = new XLFile())
            {
                file.WriteData(_data);
                file.SaveAs(basePath + ".xlsx");
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
