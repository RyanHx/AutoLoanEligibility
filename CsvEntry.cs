using CsvHelper.Configuration.Attributes;
using System;

namespace AutoLoanEligibility
{
    public class CsvEntry
    {
        [Index(0)]
        public string CompanyName { get; set; }

        [Index(1)]
        public string Address { get; set; }
    }
}
