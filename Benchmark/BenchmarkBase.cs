using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.IO;

namespace Benchmark
{
    public class BenchmarkBase
    {
        protected string _tempDirectory;
        protected const string COSNT_DUMMY_STRING = "IODJSAOIJ@OIDJASOIJONOJBOPAINEPIOQBWNI";

        [Params(10000)]
        public int NumRecords { get; set; }

        [Params(10)]
        public int NumColumnsString { get; set; }

        [Params(10)]
        public int NumColumnsNumber { get; set; }
        
        protected void Init()
        {
            _tempDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");

            if (Directory.Exists(_tempDirectory))
                new DirectoryInfo(_tempDirectory).Delete(true);

            Directory.CreateDirectory(_tempDirectory);
        }
        
        protected IEnumerable<List<object>> GetData()
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

                yield return objectlist;
            }
        }
    }
}
