using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

namespace EpPlus.PackageReader
{
    public static class ToolCreateModelFromDictionary
    {
		public static string Get(IEnumerable<string> headers, string caller)
        {
            var config = new ConfigurationBuilder()
                    .AddJsonFile("appsettings.json", true, true)
                    .Build().GetSection("CreateModelOption");
			
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(caller);
            builder.AppendLine("---------------------------------------------------------------------");
			foreach(var h in headers)
            {
                var header = h.Replace("\r", "").Replace("\n", string.Empty);
                var newHeader = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(header.ToLower());

                newHeader = newHeader
                  .Replace("][", "_");

                if (Convert.ToBoolean(config["RemoveContentInsideRoundBracket"]))
                {
                    newHeader = Regex.Replace(newHeader, @"\(.*?\)", "");
                }

				foreach(var c in config["ReplaceWithEmptyString"])
                {
                    newHeader = newHeader.Replace(c + "", string.Empty);
                }

                builder.AppendLine($"[MapHeader(\"{header}\")]");
                if (header.Contains("*"))
                {
                    builder.AppendLine("[Required]");
                }
                builder.AppendLine("public string " + newHeader + " {get; set;}");

                builder.AppendLine();
            }

            return builder.ToString();
        }
    }
}
