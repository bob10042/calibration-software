using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Pps360AmxUpc32
{
    public class Pps360AmxUpc32TestExecution
    {
        private readonly Pps360AmxUpc32Communication _communication;
        private readonly Pps360AmxUpc32Core _core;
        private readonly Pps360AmxUpc32Controls _controls;
        private readonly Pps360AmxUpc32Measurements _measurements;

        public Pps360AmxUpc32TestExecution(
            Pps360AmxUpc32Communication communication,
            Pps360AmxUpc32Core core,
            Pps360AmxUpc32Controls controls,
            Pps360AmxUpc32Measurements measurements)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
            _core = core ?? throw new ArgumentNullException(nameof(core));
            _controls = controls ?? throw new ArgumentNullException(nameof(controls));
            _measurements = measurements ?? throw new ArgumentNullException(nameof(measurements));
        }

        public void RunTest()
        {
            // opens workbook master and test worksheets          
            string mysheet = @"\\guimain\ATE\CalTestGUI\MasterPpsTest.xlsx";
            var xlApp = new Excel.Application();
            Excel.Workbooks books = xlApp.Workbooks;
            Excel.Workbook sheet = books.Open(mysheet);
            xlApp.DisplayFullScreen = true;
            xlApp.Visible = true;
            Excel.Worksheet ThreePh = xlApp.Worksheets["3 Ph"] as Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Excel.Worksheet;

            try
            {
                // front page information and start test
                FrontPage.Activate();
                ThreePh.Cells[25, 5] = _core.TestFreq;
                string date = DateTime.UtcNow.ToLongDateString();
                string time = DateTime.UtcNow.ToLongTimeString();
                FrontPage.Cells[11, 8] = time;
                FrontPage.Cells[17, 3] = date;
                FrontPage.Cells[19, 3] = "360AMX-UPC32";
                FrontPage.Cells[18, 3] = _core.Company;
                FrontPage.Cells[20, 3] = _core.SerNum;
                FrontPage.Cells[21, 3] = _core.CaseNum;
                FrontPage.Cells[17, 7] = _core.CertNum;
                FrontPage.Cells[18, 7] = _core.PlantNum;

                //three phase ac voltage test 
                ThreePh.Activate();
                _controls.Newtons4thControls();
                _controls.ThreePhaseControls();
                ExecuteThreePhaseTests(ThreePh);

                //split phase ac voltage test
                _controls.SpitPhaseAcControls();
                ExecuteSplitPhaseTests(ThreePh);

                //single phase ac voltage test
                _controls.SinglePhaseAcControls();
                ExecuteSinglePhaseTests(ThreePh);

                //frequency response
                _controls.ThreePhaseFreqRespControls();
                ExecuteFrequencyResponseTests(ThreePh);

                //reading k-factors
                var kFactors = _measurements.ReadKFactors();
                WriteKFactors(ThreePh, kFactors);

                //End off test
                FrontPage.Activate();
                time = DateTime.UtcNow.ToLongTimeString();
                FrontPage.Cells[12, 8] = time;
            }
            finally
            {
                _core.Cleanup();
            }
        }

        private void ExecuteThreePhaseTests(Excel.Worksheet ThreePh)
        {
            for (int row = 77; row <= 100; row += 3)
            {
                ExecuteThreePhaseTestSet(ThreePh, row);
            }
        }

        private void ExecuteThreePhaseTestSet(Excel.Worksheet ThreePh, int baseRow)
        {
            // Phase 1
            double voltage = Convert.ToDouble((ThreePh.Cells[baseRow, 2] as Excel.Range)?.Value);
            _communication.Write_PPS(":VOLT," + voltage);
            _core.CustomRest();
            _measurements.PpsLoopOneAcTest();
            ThreePh.Cells[baseRow, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasOne();
            ThreePh.Cells[baseRow, 8] = _measurements.GetLastMeas1().ToString();

            // Phase 2
            voltage = Convert.ToDouble((ThreePh.Cells[baseRow + 1, 2] as Excel.Range)?.Value);
            _communication.Write_PPS(":VOLT," + voltage);
            _core.CustomRest();
            _measurements.PpsLoopTwoAcTest();
            ThreePh.Cells[baseRow + 1, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasTwo();
            ThreePh.Cells[baseRow + 1, 8] = _measurements.GetLastMeas2().ToString();

            // Phase 3
            voltage = Convert.ToDouble((ThreePh.Cells[baseRow + 2, 2] as Excel.Range)?.Value);
            _communication.Write_PPS(":VOLT," + voltage);
            _core.CustomRest();
            _measurements.PpsLoopThreeAcTest();
            ThreePh.Cells[baseRow + 2, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasThree();
            ThreePh.Cells[baseRow + 2, 8] = _measurements.GetLastMeas3().ToString();
        }

        private void ExecuteSplitPhaseTests(Excel.Worksheet ThreePh)
        {
            for (int row = 126; row <= 133; row++)
            {
                double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range)?.Value);
                _communication.Write_PPS(":VOLT," + voltage);
                _core.CustomRest();
                _measurements.PpsLoopAcLineLine();
                ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
                _measurements.N4lMeasLineLine();
                ThreePh.Cells[row, 8] = _measurements.GetLastMeasLL().ToString();
            }
        }

        private void ExecuteSinglePhaseTests(Excel.Worksheet ThreePh)
        {
            for (int row = 176; row <= 183; row++)
            {
                double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range)?.Value);
                _communication.Write_PPS(":VOLT," + voltage);
                _core.CustomRest();
                _measurements.PpsLoopOneAcTest();
                ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
                _measurements.N4lMeasOne();
                ThreePh.Cells[row, 8] = _measurements.GetLastMeas1().ToString();
            }
        }

        private void ExecuteFrequencyResponseTests(Excel.Worksheet ThreePh)
        {
            for (int baseRow = 226; baseRow <= 238; baseRow += 3)
            {
                double freq = Convert.ToDouble((ThreePh.Cells[baseRow, 2] as Excel.Range)?.Value);
                _communication.Write_PPS(":FREQ," + freq);
                _core.CustomRest();
                
                _measurements.N4lMeasFour();
                ThreePh.Cells[baseRow, 4] = _measurements.GetLastMeas4().ToString();
                
                _measurements.N4lMeasOne();
                ThreePh.Cells[baseRow, 8] = _measurements.GetLastMeas1().ToString();
                
                _measurements.N4lMeasTwo();
                ThreePh.Cells[baseRow + 1, 8] = _measurements.GetLastMeas2().ToString();
                
                _measurements.N4lMeasThree();
                ThreePh.Cells[baseRow + 2, 8] = _measurements.GetLastMeas3().ToString();
            }
        }

        private void WriteKFactors(Excel.Worksheet ThreePh, 
            (string KinA, string KinB, string KinC, string KexA, string KexB, string KexC, string KiA, string KiB, string KiC, string K1P) kFactors)
        {
            ThreePh.Cells[317, 3] = kFactors.KinA;
            ThreePh.Cells[317, 6] = kFactors.KinB;
            ThreePh.Cells[317, 9] = kFactors.KinC;
            ThreePh.Cells[318, 3] = kFactors.KexA;
            ThreePh.Cells[318, 6] = kFactors.KexB;
            ThreePh.Cells[318, 9] = kFactors.KexC;
            ThreePh.Cells[319, 3] = kFactors.KiA;
            ThreePh.Cells[319, 6] = kFactors.KiB;
            ThreePh.Cells[319, 9] = kFactors.KiC;
            ThreePh.Cells[317, 10] = kFactors.K1P;
        }
    }
}
