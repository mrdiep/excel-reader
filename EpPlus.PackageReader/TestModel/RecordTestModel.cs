using System;
using System.Collections.Generic;
using System.Text;

namespace EpPlus.PackageReader.Test
{
    public class RecordTestModel
    {
        [MapHeader("[Header 2WithNew Line (km/h)]")]
        public string Header2WithNewLine { get; set; }

        [MapHeader("[Header 2][Section 1 (unit 1) *][Item 1 (unit 2/m^2) *]")]
        [Required]
        public string Header2_Section1_Item1 { get; set; }

        [MapHeader("[Header 2][Section 1 (unit 1) *][Value 1]")]
        [Required]
        public string Header2_Section1_Value1 { get; set; }

        [MapHeader("[Header 2][Section 1 (unit 1) *][Test 1]")]
        [Required]
        public string Header2_Section1_Test1 { get; set; }

        [MapHeader("[Header 2][Section 2][Item 1]")]
        public string Header2_Section2_Item1 { get; set; }

        [MapHeader("[Header 2][Section 2][Value 1]")]
        public string Header2_Section2_Value1 { get; set; }

        [MapHeader("[Header 2][Section 2][Test 1]")]
        public string Header2_Section2_Test1 { get; set; }

        [MapHeader("[Header 2][Section 3][Item 1]")]
        public string Header2_Section3_Item1 { get; set; }

        [MapHeader("[Header 2][Section 3][Value 1]")]
        public string Header2_Section3_Value1 { get; set; }

        [MapHeader("[Header 2][Section 3][Test 1]")]
        public string Header2_Section3_Test1 { get; set; }

        [MapHeader("[Header 3][Section 1]")]
        public string Header3_Section1 { get; set; }

        [MapHeader("[Header 3][Section 2]")]
        public string Header3_Section2 { get; set; }

        [MapHeader("[Header 3][Section 3][Sub 1]")]
        public string Header3_Section3_Sub1 { get; set; }

        [MapHeader("[Header 3][Section 4][Sub 2]")]
        public string Header3_Section4_Sub2 { get; set; }

        [MapHeader("[Header 4][Section 1]")]
        public string Header4_Section1 { get; set; }

        [MapHeader("[Header 4][Section 2]")]
        public string Header4_Section2 { get; set; }

        [MapHeader("[Header 4][Section 3]")]
        public string Header4_Section3 { get; set; }

        [MapHeader("[Header 4][Section 4]")]
        public string Header4_Section4 { get; set; }
    }
}
