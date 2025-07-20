using System;
using System.IO;

namespace AGXCalibrationUI.Models
{
    public class ExcelHandler
    {
        public void LoadExcelFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                // Add logic to read and parse the Excel file
                Console.WriteLine($"Excel file loaded: {filePath}");
            }
        }
    }
}