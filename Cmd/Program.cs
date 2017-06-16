using SimpleXL;
using System;
using System.Collections.Generic;
using System.IO;

namespace Cmd
{
    class Program
    {
        private const string COSNT_DUMMY_STRING = "IODJSAOIJ@OIDJASOIJONOJBOPAINEPIOQBWNI";

        static void Main(string[] args)
        {
            Console.ReadLine();
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.xlsx");
            var tempBase = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tmp");
            string basePath = Path.Combine(tempBase, Guid.NewGuid().ToString());
            using (var file = new XLFile())
            {
                /*file.ConfigureRange("A1:A2", new XLRangeConfig { Font = XLRangeFont.Bold, Border = true });
                file.ConfigureRange("B1:B2", new XLRangeConfig { Font = XLRangeFont.Bold, Border = true, Format = XLRangeFormat.Number });
                file.ConfigureRange("C1:C2", new XLRangeConfig { Format = XLRangeFormat.Percent });
                */


                file.WriteData(GetData());
                file.SaveAs(basePath + ".xlsx");
            }
            Console.ReadLine();
        }

        public static IEnumerable<List<object>> GetData()
        {
            var random = new Random();
            for (int i = 0; i < 100000; i++)
            {
                var objectlist = new List<object>();

                for (int j = 0; j < 10; j++)
                    objectlist.Add(COSNT_DUMMY_STRING + j);

                for (int j = 0; j < 10; j++)
                {
                    var x = random.Next(0, 100);
                    objectlist.Add(x);
                }

                yield return objectlist;
            }
        }
    }
}
