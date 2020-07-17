using CsvHelper;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Threading;

namespace AutoLoanEligibility
{
    class Program
    {
        #region Constants
        const string Eligible = "Eligible";
        const string Ineligible = "InEligible";
        const string Unable = "UnableToDetermine";
        #endregion

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
            ISheet eligibleSheet = workbook.CreateSheet(Eligible);
            ISheet ineligibleSheet = workbook.CreateSheet(Ineligible);
            ISheet unableSheet = workbook.CreateSheet(Unable);
            int eligibleRow = 0;
            int ineligibleRow = 0;
            int unableRow = 0;

            using (StreamReader sr = new StreamReader(csvPath))
            using (var csv = new CsvReader(sr, CultureInfo.InvariantCulture))
            {
                csv.Configuration.HasHeaderRecord = false;
                int total = 1;

                foreach (var record in csv.GetRecords<CsvEntry>())
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
                    if (json.EligibilityResult.Equals(Ineligible))
                    {
                        AddRecord(ineligibleSheet.CreateRow(ineligibleRow), record);
                        ineligibleRow++;
                    }
                    else if (json.EligibilityResult.Equals(Unable))
                    {
                        AddRecord(unableSheet.CreateRow(unableRow), record);
                        unableRow++;
                    }
                    else if (json.EligibilityResult.Equals(Eligible))
                    {
                        AddRecord(eligibleSheet.CreateRow(eligibleRow), record);
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

        private static void AddRecord(IRow row, CsvEntry record)
        {
            row.CreateCell(0).SetCellValue(record.CompanyName);
            row.CreateCell(1).SetCellValue(record.Address);
        }
    }
}
