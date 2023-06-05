using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using ConsoleApp;
using OfficeOpenXml;
using HtmlAgilityPack;



class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("This application will download all the planned Microsoft Updates for all releases");

        Console.WriteLine("Enter the path for file export example - C:\\FolderName\\Filename.xlsx");
        string filepath = Console.ReadLine();



        string apiUrl = "https://experience.dynamics.com/allreleaseplans/";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Create a new Excel package
        using (ExcelPackage package = new ExcelPackage())
        {
            // Add a new worksheet to the package
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Microsoft Releases");

            // Define the headers
            List<string> headers = new List<string>
            {
                "Product name",
                "Feature name",
                "Investment area",
                "Business value",
                "Feature details",
                "Enabled for",
                "Early access date",
                "Public preview date",
                "GA date",
                "Release wave",
                "Release Plan ID",
                "Change Description",
                "Date modified"
            };

            // Write the headers to the first row of the worksheet
            for (int i = 0; i < headers.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
            }

            int startRow = 2; // Start writing data from row 2

            int pageNumber = 1;

            Console.WriteLine("\n Download is in-progress \n");
            try
            {
                while (true)
                {
                    // Fetch data from the API
                    string apiResponse = await FetchDataFromApi(apiUrl + "?page=" + pageNumber);

                    // Deserialize the API response
                    ApiResponse response = Newtonsoft.Json.JsonConvert.DeserializeObject<ApiResponse>(apiResponse);

                    // Write the data to the worksheet
                    for (int i = 0; i < response.results.Count; i++)
                    {
                        var result = response.results[i];
                        worksheet.Cells[startRow + i, 1].Value = result["Product name"];
                        worksheet.Cells[startRow + i, 2].Value = result["Feature name"];
                        worksheet.Cells[startRow + i, 3].Value = result["Investment area"];

                        // Remove HTML tags from "Business value" field
                        string businessValue = RemoveHtmlTags(result["Business value"]);
                        worksheet.Cells[startRow + i, 4].Value = businessValue;

                        // Remove HTML tags from "Feature details" field
                        string featureDetails = RemoveHtmlTags(result["Feature details"]);
                        worksheet.Cells[startRow + i, 5].Value = featureDetails;


                        worksheet.Cells[startRow + i, 6].Value = result["Enabled for"];
                        worksheet.Cells[startRow + i, 7].Value = result["Early access date"];
                        worksheet.Cells[startRow + i, 8].Value = result["Public preview date"];
                        worksheet.Cells[startRow + i, 9].Value = result["GA date"];
                        worksheet.Cells[startRow + i, 10].Value = result["Release wave"];
                        worksheet.Cells[startRow + i, 11].Value = result["Release Plan ID"];
                        worksheet.Cells[startRow + i, 12].Value = result["Change Description"];
                        worksheet.Cells[startRow + i, 13].Value = result["Date modified"];

                    }

                    if (!response.morerecords)
                    {
                        // No more records, exit the loop
                        break;
                    }

                    // Increase the page number
                    pageNumber++;
                    startRow += response.results.Count;
                }

                // Save the Excel package to a file
                package.SaveAs(new System.IO.FileInfo(filepath));
                Console.WriteLine("Excel file created successfully. \n Press any key to close");

                Console.ReadLine();
            }
            catch(Exception ex)
            {
                Console.WriteLine(string.Format("Exception faced - {0}",ex.ToString()));
            }
           

       
        }

       
    }

    private static string RemoveHtmlTags(string html)
    {
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(html);
        return doc.DocumentNode.InnerText;
    }

    static async Task<string> FetchDataFromApi(string apiUrl)
    {
        using (HttpClient client = new HttpClient())
        {
            HttpResponseMessage response = await client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();
            string responseBody = await response.Content.ReadAsStringAsync();
            return responseBody;
        }
    }
}

