using System;

namespace SimpleXL.Attributes
{
    internal class StyleIdAttribute : Attribute
    {
        public int StyleId { get; private set; }
        public StyleIdAttribute(int styleId)
        {
            StyleId = styleId;
        }
    }
}
