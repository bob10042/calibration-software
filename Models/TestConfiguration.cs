using System;

namespace CalTestGUI.Models
{
    public class TestConfiguration
    {
        public string? VisaName { get; set; }
        public string? TestFreq { get; set; }
        public string? Company { get; set; }
        public string? CertNum { get; set; }
        public string? CaseNum { get; set; }
        public string? SerNum { get; set; }
        public string? UserFreq { get; set; }
        public string? PlantNum { get; set; }
        public DateTime TestDate { get; set; }
        public string TestTime { get; set; }

        public TestConfiguration()
        {
            TestDate = DateTime.UtcNow;
            TestTime = TestDate.ToLongTimeString();
        }
    }
}
