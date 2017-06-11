using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleXL
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
            if (!dictionary.ContainsKey(key)) {
                var count = dictionary.Keys.Count;
                dictionary.Add(key, count);
                return count;
            }

            return dictionary[key];
        }
    }
}
