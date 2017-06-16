namespace SimpleXL.Extensions
{
    internal static class IntExtensions
    {
        public static bool Between(this int number, int start, int end)
            => number >= start && number <= end;
    }
}
