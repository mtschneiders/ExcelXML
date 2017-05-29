using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelXML
{
    internal static class ConvertUtil
    {

        /// <summary>
        /// OOXML requires that "," , and &amp; be escaped, but ' and " should *not* be escaped, nor should
        /// any extended Unicode characters. This function only encodes the required characters.
        /// System.Security.SecurityElement.Escape() escapes ' and " as  &apos; and &quot;, so it cannot
        /// be used reliably. System.Web.HttpUtility.HtmlEncode overreaches as well and uses the numeric
        /// escape equivalent.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string ExcelEscapeString(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="t"></param>
        /// <param name="encodeTabCRLF"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabCRLF = false)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
                int indexAdd = 0;
                while (match.Success)
                {
                    t = t.Insert(match.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                    match = match.NextMatch();
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && ((t[i] != '\t' && t[i] != '\n' && t[i] != '\r' && encodeTabCRLF == false) || encodeTabCRLF)) //Not Tab, CR or LF
                {
                    sb.AppendFormat("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else
                {
                    sb.Append(t[i]);
                }
            }

        }
    }
}
