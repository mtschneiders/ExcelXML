using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Attributes.Jobs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Benchmark
{
    [MemoryDiagnoser]
    [ShortRunJob]
    public class EpplusBenchmark : BenchmarkBase
    {
        private List<object[]> _epplusData;

        [GlobalSetup]
        public void GlobalSetup()
        {
            Init();
            _epplusData = GetData().Select(item => item.ToArray()).ToList();
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
    }
}
