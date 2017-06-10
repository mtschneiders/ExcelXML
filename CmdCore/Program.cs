using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace CmdCore
{
    class Program
    {
        private static string _basePath;
        static void Main(string[] args)
        {
            string path = @"C:\Users\Mateus\Desktop\stuff\Book1.xlsx";
            var tempBase = Path.Combine(AppContext.BaseDirectory, "tmp");
            _basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());

            int gen0 = GC.CollectionCount(0), gen1 = GC.CollectionCount(1), gen2 = GC.CollectionCount(2);
            var stopwatch = Stopwatch.StartNew();
            using (var file = ExcelFile.LoadFromTemplate(path))
            {
                file.BeginWritingData();

                List<List<object>> data = GetData();

                foreach (var rowValues in data)
                    file.WriteRow(rowValues);

                file.EndWritingData();
                file.SaveAs(_basePath + ".xlsx");
            }
            Console.WriteLine(stopwatch.Elapsed);
            Console.WriteLine($"Gen0={GC.CollectionCount(0) - gen0} Gen1={GC.CollectionCount(1) - gen1} Gen2={GC.CollectionCount(2) - gen2}");
            Console.Read();
        }

        private static List<List<object>> GetData()
        {
            List<List<object>> data = new List<List<object>>();

            var random = new Random();
            for (int i = 0; i < 100000; i++)
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
    }
}