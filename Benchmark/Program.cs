using BenchmarkDotNet.Running;
using System;

namespace Benchmark
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Choose a benchmark to run:");
            foreach (var benchmark in (BenchmarkEnum[])Enum.GetValues(typeof(BenchmarkEnum)))
                Console.WriteLine($"{(int)benchmark} - {benchmark}");

            var input = Console.ReadLine();

            if(int.TryParse(input, out int benchmarkToRun))
            {
                Type benchmarkType = GetBenchmarkType((BenchmarkEnum)benchmarkToRun);

                if (benchmarkType != null)
                {
                    var summary = BenchmarkRunner.Run(benchmarkType);
                }
            }   
        }

        private static Type GetBenchmarkType(BenchmarkEnum benchmark)
        {
            switch (benchmark)
            {
                case BenchmarkEnum.SimpleXL:
                    return typeof(SimpleXLBenchmark);
                case BenchmarkEnum.Epplus:
                    return typeof(EpplusBenchmark);
            }

            return null;
        }

        public enum BenchmarkEnum
        {
            SimpleXL = 1,
            Epplus = 2
        }
    }
}
