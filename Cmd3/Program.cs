using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cmd3
{
    class Program
    {
        private static List<object[]> _data = GetData();

        static void Main(string[] args)
        {
            var stopwatch = Stopwatch.StartNew();
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");
            worksheet.Cell("A1").Value = _data.AsEnumerable();
            worksheet.Cell(1, 15).FormulaA1 = "K1+L1";
            workbook.SaveAs("HelloWorld.xlsx");
            stopwatch.Stop();
            Console.WriteLine(stopwatch.ElapsedMilliseconds);
            Console.Read();
        }

        private static List<object[]> GetData()
        {
            List<object[]> data = new List<object[]>();

            var random = new Random();
            for (int i = 0; i < 100000; i++)
            {
                var x = random.Next(0, 10);
                var y = Math.Round(random.NextDouble(), 2, MidpointRounding.AwayFromZero);
                data.Add(new object[]
                {
                    x,"CD BLABLA BLA BLA BLA "+i, x,"REDE BLA BLA BLA BLA"+i,x,"BRAHMA LATA 350 BLA BLA BLACX12"+i,y,y,y,y,y,y
                });
            }

            return data;
        }
    }
}
