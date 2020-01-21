using EpPlus.PackageReader.Test;
using EpPlus.PackageReader.TestModel;
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

            //var hozironItems = result.Skip(1).First().HorizontalItems.CreateInstance<Sheet2DataRecord>();
            //var verticalItems = result.Skip(1).First().VerticalItems.CreateInstance<Sheet2DataHeader>();

            foreach(var res in result)
            {
                Console.WriteLine("-----------------------" + res.SheetReaderOptionRequest.SheetName + "-----------------------");
                Console.WriteLine(ToolCreateModelFromDictionary.Get(res.HorizontalFlattenHeaders, "HorizontalFlattenHeaders"));
                Console.WriteLine(ToolCreateModelFromDictionary.Get(res.VerticalFlattenHeaders, "VerticalFlattenHeaders"));
            }

           
        }
    }
    
}
