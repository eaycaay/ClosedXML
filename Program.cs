using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        // Define the file path
        string filePath = @"C:\Users\elifa\OneDrive\Desktop\Aselsan.Aday\TestResults_deneme.xlsx";

        // Sample data
        var testResults = new List<(string Result, DateTime Time, string adı)>
        {
            ("Pass", DateTime.Now,"aa"),
            ("Fail", DateTime.Now.AddMinutes(-10),"aa"),
            ("Pass", DateTime.Now.AddMinutes(-30),"aa"),
            ("pass",DateTime.Now.AddMinutes(-40),"aa")
        };

        // Create a new workbook
        using (var workbook = new XLWorkbook())
        {
            // Add a worksheet
            var worksheet = workbook.Worksheets.Add("Results");

            // Add headers
            worksheet.Cell(1, 1).Value = "Test Result";
            worksheet.Cell(1, 2).Value = "Time";
            worksheet.Cell(1, 3).Value = "Test Cihazı";

            // Populate data
            for (int i = 0; i < testResults.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = testResults[i].Result;
                worksheet.Cell(i + 2, 2).Value = testResults[i].Time;
                worksheet.Cell(i + 2, 3).Value = testResults[i].adı;
                worksheet.Cell(i + 2, 2).Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss"; // Format date
            }

            // Auto-adjust column widths
            worksheet.Columns().AdjustToContents();

            // Save the file
            workbook.SaveAs(filePath);
            
        }

        Console.WriteLine("Excel file created successfully: " + filePath);
    }
}
