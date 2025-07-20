using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Pps3150Afx
{
    public class Pps3150AfxTestExecution
    {
        private readonly Pps3150AfxCommunication _communication;
        private readonly Pps3150AfxCore _core;
        private readonly Pps3150AfxControls _controls;
        private readonly Pps3150AfxMeasurements _measurements;

        public Pps3150AfxTestExecution(
            Pps3150AfxCommunication communication,
            Pps3150AfxCore core,
            Pps3150AfxControls controls,
            Pps3150AfxMeasurements measurements)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
            _core = core ?? throw new ArgumentNullException(nameof(core));
            _controls = controls ?? throw new ArgumentNullException(nameof(controls));
            _measurements = measurements ?? throw new ArgumentNullException(nameof(measurements));
        }

        public void RunTest()
        {
            string mysheet = @"\\guimain\ATE\CalTestGUI\MasterPpsTest.xlsx";
            var xlApp = new Excel.Application();
            Excel.Workbooks books = xlApp.Workbooks;
            Excel.Workbook sheet = books.Open(mysheet);
            xlApp.DisplayFullScreen = true;
            xlApp.Visible = true;
            Excel.Worksheet ThreePh = xlApp.Worksheets["3 Ph AFX"] as Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Excel.Worksheet;

            try
            {
                // Front page information
                FrontPage.Activate();
                ThreePh.Cells[25, 5] = _core.TestFreq;
                string date = DateTime.UtcNow.ToLongDateString();
                string time = DateTime.UtcNow.ToLongTimeString();
                FrontPage.Cells[11, 8] = time;
                FrontPage.Cells[17, 3] = date;
                FrontPage.Cells[19, 3] = "3150AFX";
                FrontPage.Cells[18, 3] = _core.Company;
                FrontPage.Cells[20, 3] = _core.SerNum;
                FrontPage.Cells[21, 3] = _core.CaseNum;
                FrontPage.Cells[17, 7] = _core.CertNum;
                FrontPage.Cells[18, 7] = _core.PlantNum;

                // Three phase AC voltage test
                ThreePh.Activate();
                _controls.Newtons4thControls();
                _controls.ThreePhaseControlsAC();
                ExecuteThreePhaseAcTests(ThreePh);

                // Frequency response test
                _controls.ThreePhaseFreqRespControls();
                ExecuteFrequencyResponseTests(ThreePh);

                // Split phase AC voltage test
                _controls.SplitPhaseAcControls();
                ExecuteSplitPhaseAcTests(ThreePh);

                // Three phase DC voltage test
                _controls.ThreePhaseControlsDC();
                ThreePh.Activate();
                _controls.Newtons4thControls();
                ExecuteThreePhaseDcTests(ThreePh);

                // Split phase DC voltage test
                _controls.SplitPhaseDcControls();
                ExecuteSplitPhaseDcTests(ThreePh);

                // Single phase tests
                if (ShowSinglePhasePrompt())
                {
                    // Single phase AC test
                    _controls.SinglePhaseAcControls();
                    ThreePh.Activate();
                    _controls.Newtons4thControls();
                    ExecuteSinglePhaseAcTests(ThreePh);

                    // Single phase DC test
                    _controls.SinglePhaseDcControls();
                    ThreePh.Activate();
                    ExecuteSinglePhaseDcTests(ThreePh);
                }

                // End of test
                FrontPage.Activate();
                time = DateTime.UtcNow.ToLongTimeString();
                FrontPage.Cells[12, 8] = time;
            }
            finally
            {
                _core.Cleanup();
            }
        }

        private void ExecuteThreePhaseAcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteThreePhaseAcTestSet(ThreePh, 77, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 80; row <= 129; row += 3)
            {
                ExecuteThreePhaseAcTestSet(ThreePh, row, 15000);
            }
        }

        private void ExecuteThreePhaseAcTestSet(Excel.Worksheet ThreePh, int baseRow, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[baseRow, 2] as Excel.Range).Value);
            _core.SetVoltageAC(voltage);
            System.Threading.Thread.Sleep(settlingTime);

            // Phase 1
            _measurements.PpsLoopOneAcTest();
            ThreePh.Cells[baseRow, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasOneAc();
            ThreePh.Cells[baseRow, 8] = _measurements.GetLastMeas1().ToString();

            // Phase 2
            _measurements.PpsLoopTwoAcTest();
            ThreePh.Cells[baseRow + 1, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasTwoAc();
            ThreePh.Cells[baseRow + 1, 8] = _measurements.GetLastMeas2().ToString();

            // Phase 3
            _measurements.PpsLoopThreeAcTest();
            ThreePh.Cells[baseRow + 2, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasThreeAc();
            ThreePh.Cells[baseRow + 2, 8] = _measurements.GetLastMeas3().ToString();
        }

        private void ExecuteFrequencyResponseTests(Excel.Worksheet ThreePh)
        {
            _core.SetVoltageAC(115);
            
            // Initial test with 30 second settling time
            ExecuteFrequencyResponseTestSet(ThreePh, 217, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 220; row <= 231; row += 3)
            {
                ExecuteFrequencyResponseTestSet(ThreePh, row, 15000);
            }
        }

        private void ExecuteFrequencyResponseTestSet(Excel.Worksheet ThreePh, int baseRow, int settlingTime)
        {
            double frequency = Convert.ToDouble((ThreePh.Cells[baseRow, 2] as Excel.Range).Value);
            _core.SetFrequency(frequency);
            System.Threading.Thread.Sleep(settlingTime);

            _measurements.N4lMeasFour();
            ThreePh.Cells[baseRow, 4] = _measurements.GetLastMeas4().ToString();
            _measurements.N4lMeasOneAc();
            ThreePh.Cells[baseRow, 8] = _measurements.GetLastMeas1().ToString();
            _measurements.N4lMeasTwoAc();
            ThreePh.Cells[baseRow + 1, 8] = _measurements.GetLastMeas2().ToString();
            _measurements.N4lMeasThreeAc();
            ThreePh.Cells[baseRow + 2, 8] = _measurements.GetLastMeas3().ToString();
        }

        private void ExecuteSplitPhaseAcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteSplitPhaseAcTest(ThreePh, 144, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 145; row <= 152; row++)
            {
                ExecuteSplitPhaseAcTest(ThreePh, row, 15000);
            }
        }

        private void ExecuteSplitPhaseAcTest(Excel.Worksheet ThreePh, int row, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range).Value);
            _core.SetVoltageAC(voltage);
            System.Threading.Thread.Sleep(settlingTime);
            _measurements.PpsLoopAcLineLine();
            ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasLineLine();
            ThreePh.Cells[row, 8] = _measurements.GetLastMeasLL().ToString();
        }

        private void ExecuteThreePhaseDcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteThreePhaseDcTestSet(ThreePh, 363, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 366; row <= 415; row += 3)
            {
                ExecuteThreePhaseDcTestSet(ThreePh, row, 15000);
            }
        }

        private void ExecuteThreePhaseDcTestSet(Excel.Worksheet ThreePh, int baseRow, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[baseRow, 2] as Excel.Range).Value);
            _core.SetVoltageDC(voltage);
            System.Threading.Thread.Sleep(settlingTime);

            // Phase 1
            _measurements.PpsLoopOneDcTest();
            ThreePh.Cells[baseRow, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasOneDc();
            ThreePh.Cells[baseRow, 8] = _measurements.GetLastMeas1().ToString();

            // Phase 2
            _measurements.PpsLoopTwoDcTest();
            ThreePh.Cells[baseRow + 1, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasTwoDc();
            ThreePh.Cells[baseRow + 1, 8] = _measurements.GetLastMeas2().ToString();

            // Phase 3
            _measurements.PpsLoopThreeDcTest();
            ThreePh.Cells[baseRow + 2, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasThreeDc();
            ThreePh.Cells[baseRow + 2, 8] = _measurements.GetLastMeas3().ToString();
        }

        private void ExecuteSplitPhaseDcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteSplitPhaseDcTest(ThreePh, 430, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 431; row <= 439; row++)
            {
                ExecuteSplitPhaseDcTest(ThreePh, row, 15000);
            }
        }

        private void ExecuteSplitPhaseDcTest(Excel.Worksheet ThreePh, int row, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range).Value);
            _core.SetVoltageDC(voltage);
            System.Threading.Thread.Sleep(settlingTime);
            _measurements.PpsLoopDcLineLine();
            ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasLineLine();
            ThreePh.Cells[row, 8] = _measurements.GetLastMeasLL().ToString();
        }

        private void ExecuteSinglePhaseAcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteSinglePhaseAcTest(ThreePh, 171, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 172; row <= 183; row++)
            {
                ExecuteSinglePhaseAcTest(ThreePh, row, 15000);
            }
        }

        private void ExecuteSinglePhaseAcTest(Excel.Worksheet ThreePh, int row, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range).Value);
            _core.SetVoltageAC(voltage);
            System.Threading.Thread.Sleep(settlingTime);
            _measurements.PpsLoopOneAcTest();
            ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasOneAc();
            ThreePh.Cells[row, 8] = _measurements.GetLastMeas1().ToString();
        }

        private void ExecuteSinglePhaseDcTests(Excel.Worksheet ThreePh)
        {
            // Initial test with 30 second settling time
            ExecuteSinglePhaseDcTest(ThreePh, 456, 30000);

            // Subsequent tests with 15 second settling time
            for (int row = 457; row <= 466; row++)
            {
                ExecuteSinglePhaseDcTest(ThreePh, row, 15000);
            }
        }

        private void ExecuteSinglePhaseDcTest(Excel.Worksheet ThreePh, int row, int settlingTime)
        {
            double voltage = Convert.ToDouble((ThreePh.Cells[row, 2] as Excel.Range).Value);
            _core.SetVoltageDC(voltage);
            System.Threading.Thread.Sleep(settlingTime);
            _measurements.PpsLoopOneDcTest();
            ThreePh.Cells[row, 4] = _measurements.GetLastPpsMeasurement().ToString();
            _measurements.N4lMeasOneDc();
            ThreePh.Cells[row, 8] = _measurements.GetLastMeas1().ToString();
        }

        private bool ShowSinglePhasePrompt()
        {
            _communication.Write_PPS(":OUTP,OFF");
            string message = "Link out the three phase output, Cancel to stop test";
            string title = "Single Phase Test";
            MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
            DialogResult result = MessageBox.Show(message, title, buttons);
            return result == DialogResult.OK;
        }
    }
}
