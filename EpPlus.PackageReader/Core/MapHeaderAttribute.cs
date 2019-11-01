using System;

namespace EpPlus.PackageReader
{
    public class MapHeaderAttribute :Attribute
    {
        public MapHeaderAttribute(string value)
        {
            Value = value;
        }

        public string Value { get; }
    }
    
}
