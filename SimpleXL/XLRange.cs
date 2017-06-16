using SimpleXL.Helpers;
using System;
using System.Text.RegularExpressions;

namespace SimpleXL
{
    internal class XLRange
    {
        public int StyleId { get; private set; }
        public string Address { get; private set; }
        public int ColumnNumberStart { get; private set; }
        public int RowNumberStart { get; private set; }
        public int ColumnNumberEnd { get; private set; }
        public int RowNumberEnd { get; private set; }

        public XLRange(string address, int styleId)
        {
            string strRegex = @"^([a-zA-Z]*)([0-9]*):([a-zA-Z]*)([0-9]*)$";
            Regex myRegex = new Regex(strRegex, RegexOptions.None);
            Match mtc = myRegex.Match(address);

            if(!mtc.Success)
                throw new ArgumentException($"Invalid range address: '{address}'");

            string columnNameStart = mtc.Groups[1].Value;
            string rowNumberStart = mtc.Groups[2].Value;
            string columnNameEnd = mtc.Groups[3].Value;
            string rowNumberEnd = mtc.Groups[4].Value;

            ColumnNumberStart = ExcelHelper.GetExcelColumnNumber(columnNameStart);
            RowNumberStart = Convert.ToInt32(rowNumberStart);
            ColumnNumberEnd = ExcelHelper.GetExcelColumnNumber(columnNameEnd);
            RowNumberEnd = Convert.ToInt32(rowNumberEnd);
            Address = address;
            StyleId = styleId;
        }

        public override int GetHashCode()
        {
            return Address.GetHashCode();
        }
    }
}
