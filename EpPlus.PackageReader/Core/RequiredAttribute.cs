using System;

namespace EpPlus.PackageReader
{
    public class RequiredAttribute : Attribute
    {
        
    }

    public class SheetNameAttribute : Attribute
    {
        public SheetNameAttribute(string value)
        {
            Value = value;
        }

        public string Value { get; }
    }

    public class HeaderRangeAttribute : Attribute
    {
        public HeaderRangeAttribute(string value, bool isHozirontalDataFlow)
        {
            Value = value;
            IsHozirontalDataFlow = isHozirontalDataFlow;
        }

        public string Value { get; }
        public bool IsHozirontalDataFlow { get; }
    }
    
}
