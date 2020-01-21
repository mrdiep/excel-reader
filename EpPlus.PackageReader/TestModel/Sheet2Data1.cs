using System;
using System.Collections.Generic;
using System.Text;

namespace EpPlus.PackageReader.TestModel
{
    class Sheet2DataHeader
    {
        [MapHeader("[Group 1][Header 1]")]
        public string Group1_Header1 { get; set; }

        [MapHeader("[Group 1][Header 2]")]
        public string Group1_Header2 { get; set; }

        [MapHeader("[Group 1][Header 3*]")]
        [Required]
        public string Group1_Header3 { get; set; }

        [MapHeader("[Group 1][Header 4]")]
        public string Group1_Header4 { get; set; }

        [MapHeader("[Group 1][Header 5]")]
        public string Group1_Header5 { get; set; }

        [MapHeader("[Group 2][Header 6]")]
        public string Group2_Header6 { get; set; }

        [MapHeader("[Group 2][Header 7]")]
        public string Group2_Header7 { get; set; }

        [MapHeader("[Group 2][Header 8]")]
        public string Group2_Header8 { get; set; }

        [MapHeader("[Group 2][Header 9]")]
        public string Group2_Header9 { get; set; }
    }

    class Sheet2DataRecord
    {
        [MapHeader("[Group 1][Header 1]")]
        public string Group1_Header1 { get; set; }

        [MapHeader("[Group 1][Header 2]")]
        public string Group1_Header2 { get; set; }

        [MapHeader("[Group 1][Header 3]")]
        public string Group1_Header3 { get; set; }

        [MapHeader("[Group 1][Header 4]")]
        public string Group1_Header4 { get; set; }

        [MapHeader("[Group 1][Header 5]")]
        public string Group1_Header5 { get; set; }

        [MapHeader("[Group 2][Header 6]")]
        public string Group2_Header6 { get; set; }

        [MapHeader("[Group 2][Header 7]")]
        public string Group2_Header7 { get; set; }

        [MapHeader("[Group 2][Header 8]")]
        public string Group2_Header8 { get; set; }

        [MapHeader("[Group 2][Header 9]")]
        public string Group2_Header9 { get; set; }
    }
}
