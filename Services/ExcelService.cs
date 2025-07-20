using System;
using System.Threading.Tasks;
using CalTestGUI.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Services
{
    public class ExcelService : IDisposable
    {
        private readonly Excel.Application xlApp;
        private readonly Excel.Workbook workbook;
        private readonly Excel.Worksheet threePh;
        private readonly Excel.Worksheet frontPage;

        public ExcelService(string excelPath)
        {
            xlApp = new Excel.Application();
            Excel.Workbooks books = xlApp.Workbooks;
            workbook = books.Open(excelPath);
            xlApp.DisplayFullScreen = true;
            xlApp.Visible = true;
            
            threePh = xlApp.Worksheets["3 Ph"] as Excel.Worksheet 
                ?? throw new InvalidOperationException("Could not find '3 Ph' worksheet");
            frontPage = xlApp.Worksheets["Front Page"] as Excel.Worksheet 
                ?? throw new InvalidOperationException("Could not find 'Front Page' worksheet");
        }

        public void InitializeFrontPage(TestConfiguration config)
        {
            frontPage.Activate();
            threePh.Cells[25, 5] = config.TestFreq;
            frontPage.Cells[11, 8] = config.TestTime;
            frontPage.Cells[17, 3] = config.TestDate.ToLongDateString();
            frontPage.Cells[19, 3] = "360ASX-UPC3";
            frontPage.Cells[18, 3] = config.Company;
            frontPage.Cells[20, 3] = config.SerNum;
            frontPage.Cells[21, 3] = config.CaseNum;
            frontPage.Cells[17, 7] = config.CertNum;
            frontPage.Cells[18, 7] = config.PlantNum;
        }

        public double? GetCellValue(int row, int col)
        {
            var range = threePh.Cells[row, col] as Excel.Range;
            return range?.Value2 as double?;
        }

        public void SetMeasurement(int row, int col, double value)
        {
            threePh.Cells[row, col] = value.ToString();
        }

        public void SetTestEndTime(string time)
        {
            frontPage.Activate();
            frontPage.Cells[12, 8] = time;
        }

        public void Dispose()
        {
            workbook?.Close();
            xlApp?.Quit();
        }
    }
}
