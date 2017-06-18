using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Attributes.Jobs;
using SimpleXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Benchmark
{
    [MemoryDiagnoser]
    [ShortRunJob]
    public class SimpleXLBenchmark : BenchmarkBase
    {
        private List<List<object>> _simpleXLData;

        [GlobalSetup]
        public void GlobalSetup()
        {
            Init();
            _simpleXLData = GetData().ToList();
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
    }
}
