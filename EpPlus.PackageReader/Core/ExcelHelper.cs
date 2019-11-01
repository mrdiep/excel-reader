using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EpPlus.PackageReader
{
    public class ExcelHelper
    {
        public IEnumerable<SheetReaderResponse> ReadExcel(Stream excelStream, params SheetReaderOptions[] options)
        {
            var list = new List<SheetReaderResponse>();
            using (var excelPackage = new ExcelPackage(excelStream))
            {
                var worksheets = excelPackage.Workbook.Worksheets;
                foreach(var option in options)
                {
                    IEnumerable<string> horizontalFlattenHeaders;
                    IEnumerable< string > verticalFlattenHeaders;

                    var response = new SheetReaderResponse()
                    {
                        SheetReaderOptionRequest = option,
                        HorizontalItems = GetHorizontalItems(worksheets[option.SheetName], option, out horizontalFlattenHeaders),
                        VerticalItems = GetVerticalItems(worksheets[option.SheetName], option, out verticalFlattenHeaders)
                    };

                    response.VerticalFlattenHeaders = verticalFlattenHeaders;
                    response.HorizontalFlattenHeaders = horizontalFlattenHeaders;

                    list.Add(response);
                }
            }

            return list;
        }

        private string GetMergedRangeAddress(ExcelRange excelRange)
        {
            if (excelRange.Merge)
            {
                var idx = excelRange.Worksheet.GetMergeCellId(excelRange.Start.Row, excelRange.Start.Column);
                var cells = excelRange.Worksheet.MergedCells[idx - 1].Split(":", StringSplitOptions.RemoveEmptyEntries);
                return cells[0];
            }
            else
            {
                return excelRange.Address;
            }
        }

        private IEnumerable<Dictionary<string, object>> GetHorizontalItems(ExcelWorksheet worksheet, SheetReaderOptions option, out IEnumerable<string> horizontalFlattenHeaders)
        {
            var headerDict = new Dictionary<string, string>();
            var headers = worksheet.Cells[option.HorizontalHeaderExcelRange];
            for(var col = headers.Start.Column; col <= headers.End.Column; col++)
            {
                var list = new List<string>();
                for (var row = headers.Start.Row; row <= headers.End.Row; row++)
                {
                    var address = GetMergedRangeAddress(worksheet.Cells[row, col]);
                    var header = worksheet.Cells[address].Text;
                    header = header?.Replace("\r", string.Empty).Replace("\n", string.Empty);
                    list.Add($@"[{header}]");
                }

                var columName = Regex.Replace(worksheet.Cells[1, col].Address, "\\d+$", string.Empty);

                headerDict.Add(columName, string.Join("", list.Where(x => !string.IsNullOrWhiteSpace(x) && x!="[]").Distinct()));
            }
            horizontalFlattenHeaders = headerDict.Values;

            var rowItems = new List<Dictionary<string, object>>();
            for (var row = headers.End.Row + 1; row <= worksheet.Dimension.End.Row; row++)
            {
                var item = new Dictionary<string, object>();
                for (int col = headers.Start.Column; col <= headers.End.Column; col++)
                {
                    if (worksheet.Cells[row, col].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                    {
                        var columName = Regex.Replace(worksheet.Cells[row, col].Address, "\\d+$", string.Empty);
                        item.Add(headerDict[columName], worksheet.Cells[row, col].Value);
                    }
                }

                if (item.Any() && !item.Values.All(x => x == null))
                {
                    rowItems.Add(item);
                }
            }

            return rowItems;
        }

        private IEnumerable<Dictionary<string, object>> GetVerticalItems(ExcelWorksheet worksheet, SheetReaderOptions option, out IEnumerable<string> verticalFlattenHeaders)
        {
            var headerDict = new Dictionary<int, string>();
            var headers = worksheet.Cells[option.VerticalHeaderExcelRange];
            for (var row = headers.Start.Row; row <= headers.End.Row; row++)
            {
                var list = new List<string>();
                for (var col = headers.Start.Column; col <= headers.End.Column; col++)
                {
                    var address = GetMergedRangeAddress(worksheet.Cells[row, col]);
                    var header = worksheet.Cells[address].Text;
                    header = header?.Replace("\r", string.Empty).Replace("\n", string.Empty);
                    list.Add($@"[{header}]");
                }

                headerDict.Add(row, string.Join("", list.Where(x => !string.IsNullOrWhiteSpace(x) && x != "[]").Distinct()));
            }
            verticalFlattenHeaders = headerDict.Values;

            var rowItems = new List<Dictionary<string, object>>();
            for (var col = headers.End.Column + 1; col <= worksheet.Dimension.End.Column; col++)
            {
                var item = new Dictionary<string, object>();
                for (int row = headers.Start.Row; row <= headers.End.Row; row++)
                {
                    if (worksheet.Cells[row, col].Value != null && !string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                    {
                        item.Add(headerDict[row], worksheet.Cells[row, col].Value);
                    }
                }

                if (item.Any() && !item.Values.All(x => x == null))
                {
                    rowItems.Add(item);
                }
            }

            return rowItems;
        }

    }

    
}
