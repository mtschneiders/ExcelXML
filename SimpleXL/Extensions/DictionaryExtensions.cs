using System.Collections.Generic;

namespace ExcelXML.Extensions
{
    internal static class DictionaryExtensions
    {
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
    }
}
