using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Pps360AsxUpc3
{
    public partial class Pps360AsxUpc3Core
    {
        public void RunTest()
        {
            try
            {
                // opens workbook master and test worksheets          
                string mysheet = @"\\guimain\ATE\CalTestGUI\MasterPpsTest.xlsx";
                var xlApp = new Excel.Application();
                Excel.Workbooks books = xlApp.Workbooks;
                Excel.Workbook sheet = books.Open(mysheet);
                xlApp.DisplayFullScreen = true;
                xlApp.Visible = true;
                Excel.Worksheet? ThreePh = xlApp.Worksheets["3 Ph"] as Excel.Worksheet;
                Excel.Worksheet? FrontPage = xlApp.Worksheets["Front Page"] as Excel.Worksheet;

                if (ThreePh == null || FrontPage == null)
                {
                    MessageBox.Show("Could not open required Excel worksheets");
                    return;
                }

                try
                {
                    // Front page information
                    FrontPage.Activate();
                    ThreePh.Cells[25, 5] = TestFreq;
                    date = DateTime.UtcNow.ToLongDateString();
                    time = DateTime.UtcNow.ToLongTimeString();
                    FrontPage.Cells[11, 8] = time;
                    FrontPage.Cells[17, 3] = date;
                    FrontPage.Cells[19, 3] = "360ASX-UPC3";
                    FrontPage.Cells[18, 3] = Company;
                    FrontPage.Cells[20, 3] = serNum;
                    FrontPage.Cells[21, 3] = CaseNum;
                    FrontPage.Cells[17, 7] = CertNum;
                    FrontPage.Cells[18, 7] = PlantNum;

                    // Three phase AC voltage test
                    ThreePh.Activate();
                    Newtons4thControls();
                    ThreePhaseControls();

                    // Execute voltage measurements for different ranges
                    ExecuteVoltageRange(ThreePh, 77, 79);  // First range
                    ExecuteVoltageRange(ThreePh, 80, 82);  // Second range
                    ExecuteVoltageRange(ThreePh, 83, 85);  // Third range
                    ExecuteVoltageRange(ThreePh, 86, 88);  // Fourth range
                    ExecuteVoltageRange(ThreePh, 89, 91);  // Fifth range
                    ExecuteVoltageRange(ThreePh, 92, 94);  // Sixth range
                    ExecuteVoltageRange(ThreePh, 95, 97);  // Seventh range
                    ExecuteVoltageRange(ThreePh, 98, 100); // Eighth range

                    // Split phase AC voltage test
                    SpitPhaseAcControls();
                    ExecuteLineLineTests(ThreePh);

                    // Single phase AC voltage test
                    SinglePhaseAcControls();
                    ExecuteSinglePhaseTests(ThreePh);

                    // Frequency response test
                    ThreePhaseFreqRespControls();
                    ExecuteFrequencyResponseTests(ThreePh);
                }
                finally
                {
                    // Cleanup if needed
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during test execution: {ex.Message}");
            }
        }

        private void ExecuteVoltageRange(Excel.Worksheet sheet, int startRow, int endRow)
        {
            var range = sheet.Cells[startRow, 2] as Excel.Range;
            if (range?.Value2 is double value)
            {
                cellvalue = value;
                Write_PPS(":VOLT," + cellvalue);
            }

            System.Threading.Thread.Sleep(startRow == 77 ? 30000 : 15000); // Longer delay for first range

            // Phase 1
            PpsLoopOneAcTest();
            sheet.Cells[startRow, 4] = pps1.ToString();
            N4lMeasOne();
            sheet.Cells[startRow, 8] = meas1.ToString();

            // Phase 2
            PpsLoopTwoAcTest();
            sheet.Cells[startRow + 1, 4] = pps1.ToString();
            N4lMeasTwo();
            sheet.Cells[startRow + 1, 8] = meas2.ToString();

            // Phase 3
            PpsLoopThreeAcTest();
            sheet.Cells[startRow + 2, 4] = pps1.ToString();
            N4lMeasThree();
            sheet.Cells[startRow + 2, 8] = meas3.ToString();
        }

        private void ExecuteLineLineTests(Excel.Worksheet sheet)
        {
            // Initial setup
            var range = sheet.Cells[126, 2] as Excel.Range;
            if (range?.Value2 is double value)
            {
                cellvalue = value;
                Write_PPS(":VOLT," + cellvalue);
            }

            System.Threading.Thread.Sleep(30000);
            PpsLoopAcLineLine();
            sheet.Cells[126, 4] = pps1.ToString();
            N4lMeasLineLine();
            sheet.Cells[126, 8] = measlL.ToString();

            // Subsequent ranges
            for (int i = 127; i <= 133; i++)
            {
                range = sheet.Cells[i, 2] as Excel.Range;
                if (range?.Value2 is double val)
                {
                    cellvalue = val;
                    Write_PPS(":VOLT," + cellvalue);
                }
                System.Threading.Thread.Sleep(15000);
                PpsLoopAcLineLine();
                sheet.Cells[i, 4] = pps1.ToString();
                N4lMeasLineLine();
                sheet.Cells[i, 8] = measlL.ToString();
            }
        }

        private void ExecuteSinglePhaseTests(Excel.Worksheet sheet)
        {
            // Initial setup
            var range = sheet.Cells[176, 2] as Excel.Range;
            if (range?.Value2 is double value)
            {
                cellvalue = value;
                Write_PPS(":VOLT," + cellvalue);
            }

            System.Threading.Thread.Sleep(30000);
            PpsLoopOneAcTest();
            sheet.Cells[176, 4] = pps1.ToString();
            N4lMeasOne();
            sheet.Cells[176, 8] = meas1.ToString();

            // Subsequent ranges
            for (int i = 177; i <= 183; i++)
            {
                range = sheet.Cells[i, 2] as Excel.Range;
                if (range?.Value2 is double val)
                {
                    cellvalue = val;
                    Write_PPS(":VOLT," + cellvalue);
                }
                System.Threading.Thread.Sleep(15000);
                PpsLoopOneAcTest();
                sheet.Cells[i, 4] = pps1.ToString();
                N4lMeasOne();
                sheet.Cells[i, 8] = meas1.ToString();
            }
        }

        private void ExecuteFrequencyResponseTests(Excel.Worksheet sheet)
        {
            // First frequency test
            var range = sheet.Cells[226, 2] as Excel.Range;
            if (range?.Value2 is double value)
            {
                cellvalue = value;
                Write_PPS(":FREQ," + cellvalue);
            }

            System.Threading.Thread.Sleep(30000);
            N4lMeasFour();
            sheet.Cells[226, 4] = meas4.ToString();
            N4lMeasOne();
            sheet.Cells[226, 8] = meas1.ToString();
            N4lMeasTwo();
            sheet.Cells[227, 8] = meas2.ToString();
            N4lMeasThree();
            sheet.Cells[228, 8] = meas3.ToString();

            // Second frequency test
            range = sheet.Cells[229, 2] as Excel.Range;
            if (range?.Value2 is double val229)
            {
                cellvalue = val229;
                Write_PPS(":FREQ," + cellvalue);
            }

            System.Threading.Thread.Sleep(15000);
            N4lMeasFour();
            sheet.Cells[229, 4] = meas4.ToString();
            N4lMeasOne();
            sheet.Cells[229, 8] = meas1.ToString();
            N4lMeasTwo();
            sheet.Cells[230, 8] = meas2.ToString();
            N4lMeasThree();
            sheet.Cells[231, 8] = meas3.ToString();

            // Third frequency test
            range = sheet.Cells[232, 2] as Excel.Range;
            if (range?.Value2 is double val232)
            {
                cellvalue = val232;
                Write_PPS(":FREQ," + cellvalue);
            }

            System.Threading.Thread.Sleep(15000);
            N4lMeasFour();
            sheet.Cells[232, 4] = meas4.ToString();
            N4lMeasOne();
            sheet.Cells[232, 8] = meas1.ToString();
            N4lMeasTwo();
            sheet.Cells[233, 8] = meas2.ToString();
            N4lMeasThree();
            sheet.Cells[234, 8] = meas3.ToString();
        }
    }
}
