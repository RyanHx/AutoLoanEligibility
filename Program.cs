using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using System;
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
            {
                var address = sr.ReadLine();
                int eligibleRow = 0;
                int ineligibleRow = 0;
                int unableRow = 0;
                int total = 1;

                while (address != null)
                {
                    Console.WriteLine($"Checking address {total}");
                    if (address.StartsWith('"'))
                    {
                        address = address.Substring(1, address.Length - 2);
                    }

                    Thread.Sleep(TimeSpan.FromSeconds(2));
                    HttpResponseMessage response = client.GetAsync("https://eligibility.sc.egov.usda.gov/eligibility/MapAddressVerification?address=" + address + "&whichapp=RBSIELG").GetAwaiter().GetResult();
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Connection error. Too many requests?");
                        Console.WriteLine("Writing current results...");
                        break;
                    }

                    Temperatures json = JsonConvert.DeserializeObject<Temperatures>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                    if (json.EligibilityResult.Equals("InEligible"))
                    {
                        ineligibleSheet.CreateRow(ineligibleRow).CreateCell(0).SetCellValue(address);
                        ineligibleRow++;
                    }
                    else if (json.EligibilityResult.Equals("UnableToDetermine"))
                    {
                        unableSheet.CreateRow(unableRow).CreateCell(0).SetCellValue(address);
                        unableRow++;
                    }
                    else
                    {
                        eligibleSheet.CreateRow(eligibleRow).CreateCell(0).SetCellValue(address);
                        eligibleRow++;
                    }
                    address = sr.ReadLine();
                    total++;
                }

            }
            FileStream fs = File.Create("output.xlsx");
            workbook.Write(fs);
            fs.Close();
            Console.WriteLine("Task complete");
            Console.WriteLine($"Output directory: {AppDomain.CurrentDomain.BaseDirectory}output.xlsx");
            Console.WriteLine("Press any key to close this window...");
            Console.ReadKey();
        }
    }
}
