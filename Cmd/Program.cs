using ExcelXML;
using System;
using System.Collections.Generic;
using System.IO;

namespace Cmd
{
    class Program
    {
        private const string COSNT_RANDOM_STRING = "IODJSAOIJ@OIDJASOIJONOJBOPAINEPIOQBWNI";

        static void Main(string[] args)
        {
            Console.ReadLine();
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.xlsx");
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            string basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());
            using (var file = ExcelFile.LoadFromTemplate(path))
            {
                file.BeginWritingData();
                
                foreach (var rowValues in GetData())
                    file.WriteRow(rowValues);

                file.EndWritingData();
                file.SaveAs(basePath + ".xlsx");
            }

            Console.ReadLine();
        }

        public static IEnumerable<List<object>> GetData()
        {
            var random = new Random();
            for (int i = 0; i < 10000; i++)
            {
                int lineNumber = i + 3;
                var x = random.Next(0, 10);
                var y = Math.Round(random.NextDouble(), 2, MidpointRounding.AwayFromZero);
                yield return new List<object>
                {
                    x, COSNT_RANDOM_STRING+i, x, COSNT_RANDOM_STRING+i, x, COSNT_RANDOM_STRING+i, y, y, y, y, y, y, $"=(I{lineNumber}+J{lineNumber})*(H{lineNumber}/100)"
                };
            }
        }
    }
}
