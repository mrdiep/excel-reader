using EpPlus.PackageReader.Test;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;

namespace EpPlus.PackageReader
{
    class Program
    {
        static void Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                    .AddJsonFile("appsettings.json", true, true)
                    .Build();

            var fileInfo = new FileInfo(config["ExcelFilePath"]);
            var result = new ExcelHelper().ReadExcel(fileInfo.OpenRead(), config.GetSection("SheetReaderOptions").GetChildren().Select(x => new SheetReaderOptions {
                HorizontalHeaderExcelRange = x["HorizontalHeaderExcelRange"],
                SheetName = x["SheetName"],
                VerticalHeaderExcelRange = x["VerticalHeaderExcelRange"],
            }).ToArray());

            //var hozironItems = result.First().HorizontalItems.CreateInstance<RecordTestModel>();
            //var verticalItems = result.First().VerticalItems.CreateInstance<HeaderTestModel>();

            foreach(var res in result)
            {
                Console.WriteLine("-----------------------" + res.SheetReaderOptionRequest.SheetName + "-----------------------");
                Console.WriteLine(ToolCreateModelFromDictionary.Get(res.HorizontalFlattenHeaders, "HorizontalFlattenHeaders"));
                Console.WriteLine(ToolCreateModelFromDictionary.Get(res.VerticalFlattenHeaders, "VerticalFlattenHeaders"));
            }

           
        }
    }
    
}
