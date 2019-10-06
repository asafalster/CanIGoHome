using System.Collections.Generic;
using System.Collections.Specialized;

namespace CanIGoHome
{
    public static partial class Extensions
    {
        /// <summary>
        ///     A NameValueCollection extension method that converts the @this to a dictionary.
        /// </summary>
        /// <param name="this">The @this to act on.</param>
        /// <returns>@this as an IDictionary&lt;string,object&gt;</returns>
        public static IDictionary<string, string> ToDictionary(this NameValueCollection @this, IDictionary<string, string> sourceDict = null)
        {
            var dict = sourceDict ?? new Dictionary<string, string>();

            if (@this != null)
            {
                foreach (string key in @this.AllKeys)
                {
                    dict.Add(key, @this[key]);
                }
            }

            return dict;
        }
    }
}
