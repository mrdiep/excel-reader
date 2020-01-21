using System;
using System.Collections.Generic;
using System.Linq;

namespace EpPlus.PackageReader
{
    public class SheetReaderResponse
    {
        public IEnumerable<string> HorizontalFlattenHeaders { get; set; }

        public IEnumerable<string> VerticalFlattenHeaders { get; set; }

        public IEnumerable<Dictionary<string, object>> HorizontalItems { get; set; }

        public IEnumerable<Dictionary<string, object>> VerticalItems { get; set; }

        public SheetReaderOptions SheetReaderOptionRequest { get; set; }
    }

    public static class SheetReaderResponseExtenstion
    {
        public static T[] CreateInstance<T>(this IEnumerable<Dictionary<string, object>> rowData)
        {
            List<T> items = new List<T>();
            var type = typeof(T);
            var headerRef = type.GetProperties()
                .Select(x => new {
                    x.Name,
                    Header = x.GetCustomAttributes(typeof(MapHeaderAttribute), false).Cast<MapHeaderAttribute>().First().Value
                })
                .ToDictionary(x => x.Header, x => x.Name);

            foreach(var rowItem in rowData)
            {
                var t = Activator.CreateInstance<T>();

                foreach(var prop in rowItem)
                {
                    var propRef = type.GetProperty(headerRef[prop.Key]);
                    propRef.SetValue(t, Convert.ChangeType(prop.Value, propRef.PropertyType));
                }
                items.Add(t);
            }

            return items.ToArray();
        }
    }
}
