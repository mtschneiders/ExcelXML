using System;

namespace SimpleXL.Extensions
{
    internal static class DateTimeExtensions
    {
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
