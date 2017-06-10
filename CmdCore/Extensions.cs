using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CmdCore
{
    internal static class Extensions
    {
        public static int IndexOfOrNew(this List<string> list, string value)
        {
            int index = list.IndexOf(value);
            if (index == -1)
            {
                list.Add(value);
                index = list.Count - 1;
            }
            return index;
        }

        public static int GetValueOrNew<T1>(this Dictionary<T1, int> dictionary, T1 key)
        {
            if (!dictionary.ContainsKey(key))
            {
                var count = dictionary.Keys.Count;
                dictionary.Add(key, count);
                return count;
            }

            return dictionary[key];
        }

        public static double ToOADate(this DateTime date)
        {
            long value = date.Ticks;
            if (value == 0L)
            {
                return 0.0;
            }
            if (value < 864000000000L)
            {
                value += 599264352000000000L;
            }
            if (value < 31241376000000000L)
            {
                throw new OverflowException();
            }
            long num = (value - 599264352000000000L) / 10000L;
            if (num < 0L)
            {
                long num2 = num % 86400000L;
                if (num2 != 0L)
                {
                    num -= (86400000L + num2) * 2L;
                }
            }

            return (double)num / 86400000.0;            
        }


    }
}
