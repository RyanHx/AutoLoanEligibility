using CsvHelper;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Threading;

namespace AutoLoanEligibility
{
    class Program
    {
        static HttpClient client = new HttpClient();
        static void Main(string[] args)
        {
            string csvPath = "";
            while (!File.Exists(csvPath))
            {
                Console.WriteLine("Enter path to CSV (e.g. C:\\Desktop\\test.csv): ");
                csvPath = Console.ReadLine();
            }

            XSSFWorkbook workbook = new XSSFWorkbook();
            var eligibleSheet = workbook.CreateSheet("Eligible");
            var ineligibleSheet = workbook.CreateSheet("Ineligible");
            var unableSheet = workbook.CreateSheet("UnableToDetermine");

            using (StreamReader sr = new StreamReader(csvPath))
            using (var csv = new CsvReader(sr, CultureInfo.InvariantCulture))
            {
                csv.Configuration.HasHeaderRecord = false;
                var records = csv.GetRecords<CsvEntry>();
                #region Excel output row numbers
                int eligibleRow = 0;
                int ineligibleRow = 0;
                int unableRow = 0;
                int total = 1;
                #endregion

                foreach (var record in records)
                {
                    Console.WriteLine($"Checking address {total}");
                    Thread.Sleep(TimeSpan.FromSeconds(2)); // Prevent spamming requests
                    HttpResponseMessage response = client.GetAsync("https://eligibility.sc.egov.usda.gov/eligibility/MapAddressVerification?address=" + record.Address + "&whichapp=RBSIELG").GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Connection error:\nStatus code: {response.StatusCode}\nMessage: {response.ReasonPhrase}");
                        Console.WriteLine($"Failed on: {record.Address}");
                        Console.WriteLine("Writing current results...");
                        break;
                    }

                    EligibilityModel json = JsonConvert.DeserializeObject<EligibilityModel>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    if (json.EligibilityResult.Equals("InEligible"))
                    {
                        var row = ineligibleSheet.CreateRow(ineligibleRow);
                        row.CreateCell(0).SetCellValue(record.CompanyName);
                        row.CreateCell(1).SetCellValue(record.Address);
                        ineligibleRow++;
                    }
                    else if (json.EligibilityResult.Equals("UnableToDetermine"))
                    {
                        var row = unableSheet.CreateRow(unableRow);
                        row.CreateCell(0).SetCellValue(record.CompanyName);
                        row.CreateCell(1).SetCellValue(record.Address);
                        unableRow++;
                    }
                    else
                    {
                        var row = eligibleSheet.CreateRow(eligibleRow);
                        row.CreateCell(0).SetCellValue(record.CompanyName);
                        row.CreateCell(1).SetCellValue(record.Address);
                        eligibleRow++;
                    }
                    total++;
                }
            }          
            FileStream fs = File.Create("output.xlsx");
            workbook.Write(fs);
            fs.Close();
            Console.WriteLine("Task complete");
            Console.WriteLine($"Output directory: {Directory.GetCurrentDirectory()}\\output.xlsx");
            Console.WriteLine("Press any key to close this window...");
            Console.ReadKey();
        }
    }
}
