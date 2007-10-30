using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace GratisInc.Tools.FogBugz.WorkingOn
{
    public static class Extensions
    {
        /// <summary>
        /// Returns a Dictionary containing the KeyValuePairs in the specified collection.
        /// </summary>
        /// <typeparam name="TKey">The type of the Key.</typeparam>
        /// <typeparam name="TValue">The type of the Value.</typeparam>
        /// <param name="col">A collection of KeyValuePairs.</param>
        /// <returns>A dictionary consisting of the KeyValuePairs from col.</returns>
        public static Dictionary<TKey, TValue> FromKeyValuePairCollection<TKey, TValue>(this Dictionary<TKey, TValue> dict, IEnumerable<KeyValuePair<TKey, TValue>> col)
        {
            dict = new Dictionary<TKey, TValue>();
            foreach (KeyValuePair<TKey, TValue> kvp in col)
            {
                dict.Add(kvp.Key, kvp.Value);
            }
            return dict;
        }

        /// <summary>
        /// Evaluates the result of a FogBugz API call and determines if the result is an error.
        /// </summary>
        /// <param name="error">The resulting FogBugz error, if any.</param>
        /// <returns>A boolean value representing whether the FogBugz API call was an error</returns>
        public static Boolean IsFogBugzError(this XDocument doc, out FogBugzApiError error)
        {
            if (doc.Descendants("error").Count() == 1)
            {
                error = 
                    (
                    from c in doc.Descendants("error")
                    select new FogBugzApiError
                    {
                        Code = Int32.Parse(c.Attribute("code").Value),
                        Message = c.Value
                    }
                    ).First();
                return true;
            }
            else
            {
                error = null;
                return false;
            }
        }
    }
}
