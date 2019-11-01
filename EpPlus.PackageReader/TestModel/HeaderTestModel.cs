using System;
using System.Collections.Generic;
using System.Text;

namespace EpPlus.PackageReader.Test
{
    public class HeaderTestModel
    {
        [MapHeader("[Hoziron Header 1][A1]")]
        public string HozironHeader1_A1 { get; set; }

        [MapHeader("[Hoziron Header 1][A2]")]
        public string HozironHeader1_A2 { get; set; }

        [MapHeader("[Hoziron Header 2][A3]")]
        public string HozironHeader2_A3 { get; set; }

        [MapHeader("[Hoziron Header 2][A4]")]
        public string HozironHeader2_A4 { get; set; }

        [MapHeader("[Start Date]")]
        public string StartDate { get; set; }
    }
}
