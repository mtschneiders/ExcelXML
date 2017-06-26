namespace SimpleXL
{
    /// <summary> Represents a set of configuration properties for an excel range
    /// </summary>
    public class XLRangeConfig
    {
        /// <summary> Indicate if the range will have default border applied
        /// </summary>
        public bool Border { get; set; }

        /// <summary> Defines Range format
        /// </summary>
        public XLRangeFormat Format { get; set; }

        /// <summary> Defines Font Style
        /// </summary>
        public XLRangeFont Font { get; set; }
    }
}
