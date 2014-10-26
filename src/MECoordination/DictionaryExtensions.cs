using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalToolSuite.MECoordination
{
    public static class DictionaryExtensions
    {
        public static Dictionary<TKey, TValue> ToDictionary<TKey, TValue>(this IEnumerable<KeyValuePair<TKey, TValue>> source)
        {
            return source.ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        public static Dictionary<TKey, TValue> Where<TKey, TValue>(this IDictionary<TKey, TValue> source, Func<KeyValuePair<TKey, TValue>, bool> predicate)
        {
            return Enumerable.Where(source, predicate).ToDictionary(kv => kv.Key, kv => kv.Value);
        }
    }
}
