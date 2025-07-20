using NationalInstruments.Visa;
using Ivi.Visa;
using System;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI
{
    /*
     External class to test the Pacific 118ACX_UPC1. 
     All communications and external equipment is set
     up through this case. 
    */
    class Pps_118AcxUpc1
    {
        private readonly SerialPort SerWriteN4l = new SerialPort();
        public MessageBasedSession? mbSession = null;
        
        // Properties initialized through constructor
        public string visaName { get; set; }
        public string TestFreq { get; set; }
        public string Company { get; set; }
        public string CertNum { get; set; }
        public string CaseNum { get; set; }
        public string serNum { get; set; }
        public string UserFreq { get; set; }
        public string PlantNum { get; set; }

        private string? date = null;
        private string? time = null;
        private string? PhaseA = null;
        private string? PhaseB = null;
        private string? PhaseC = null;
        private string? Frequency = null;
        private string? PhaseAb = null;

        private double meas1 = 0;
        private double meas2 = 0;
        private double meas3 = 0;
        private double meas4 = 0;
        private double measlL = 0;
        private double pps1 = 0;
        private double cellvalue = 0;

        public Pps_118AcxUpc1(string visaName, string testFreq, string company, string certNum, 
                             string caseNum, string serNum, string userFreq, string plantNum)
        {
            this.visaName = visaName;
            this.TestFreq = testFreq;
            this.Company = company;
            this.CertNum = certNum;
            this.CaseNum = caseNum;
            this.serNum = serNum;
            this.UserFreq = userFreq;
            this.PlantNum = plantNum;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
        }

        private string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        private string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "\\n").Replace("\r", "\\r");
        }

        public void OpenSession()
        {
            if (string.IsNullOrEmpty(visaName))
            {
                MessageBox.Show("Select Visa Address");
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            using (var rmSession = new ResourceManager())
            {
                try
                {
                    mbSession = (MessageBasedSession)rmSession.Open(visaName);
                    mbSession.TimeoutMilliseconds = 5000;
                    mbSession.TerminationCharacter = 0xA;
                    mbSession.SendEndEnabled = false;
                    mbSession.TerminationCharacterEnabled = true;
                }
                catch (InvalidCastException)
                {
                    MessageBox.Show("Resource selected must be a message-based session");
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message);
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
            Write_PPS("*IDN?");
            Read_Pps();
            Thread.Sleep(1000);
        }

        private void Write_PPS(string command)
        {
            if (mbSession == null) return;
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        private string Read_Pps()
        {
            try
            {
                if (mbSession == null) return string.Empty;
                string messageP = InsertCommonEscapeSequences((mbSession.RawIO.ReadString()));
                return messageP.Replace(@"\n", "").Replace(@"\r", "").Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }

        public void ClosePort()
        {
            if (SerWriteN4l.IsOpen)
            {
                SerWriteN4l.Close();
            }
            if (mbSession != null)
            {
                mbSession.Dispose();
            }
        }

        public string PortName_Newtowns_4th
        {
            get => SerWriteN4l.PortName;
            set => SerWriteN4l.PortName = value;
        }

        public bool IsOpen_Newtowns_4th => SerWriteN4l.IsOpen;

        public void Open_Newtowns_4th()
        {
            SerWriteN4l.BaudRate = 9600;
            SerWriteN4l.Parity = Parity.None;
            SerWriteN4l.StopBits = StopBits.One;
            SerWriteN4l.DataBits = 8;
            SerWriteN4l.Handshake = Handshake.None;
            SerWriteN4l.RtsEnable = true;
            SerWriteN4l.DtrEnable = true;
            SerWriteN4l.WriteTimeout = 2000;
            SerWriteN4l.ReadTimeout = 2000;
            SerWriteN4l.Open();
        }

        private void WriteN4l(string command)
        {
            SerWriteN4l.WriteLine(command);
            Thread.Sleep(200);
        }

        private string Read_N4l()
        {
            try
            {
                string messageN = SerWriteN4l.ReadLine();
                return messageN.Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }

        public void Close()
        {
            SerWriteN4l.Close();
        }

        private void CustomRest()
        {
            Thread.Sleep(2000);
        }

        private void Newtons4thControls()
        {
            WriteN4l("*RST");
            WriteN4l("TRG");                      // initiates a new measurement, resets the range and smoothing
            WriteN4l("KEYBOARD,DISABLE");         // disable keyboard
            WriteN4l("SYST;POWER");               // selects power meter
            WriteN4l("COUPLI,PHASE1,AC+DC");      // coupling ac + dc
            WriteN4l("COUPLI,PHASE2,AC+DC");      // coupling ac + dc
            WriteN4l("COUPLI,PHASE3,AC+DC");      // coupling ac + dc
            WriteN4l("APPLIC,NORMAL,SPEED3");     // select application mode
            WriteN4l("BANDWI,WIDE");              // bandwidth @ 3MHz
            WriteN4l("DATALO,DISABLE");           // datalog disabled
            WriteN4l("RESOLU,HIGH");              // sets the data resolution                             
            WriteN4l("EFFICI,0");                 // efficiency calculation disabled
            WriteN4l("FAST,ON");                  // this will disable the screen for faster resolutions 
            WriteN4l("FREQFI,OFF");               // turns the frequency filter off                     
            WriteN4l("SPEED,HIGH");               // slows the speed down 
            WriteN4l("DISPLAY,VOLTAGE");          // display voltage mode
            WriteN4l("ZOOM,1,3,5,6,7");           // sets the zoom level for the display
            Thread.Sleep(5000);
            WriteN4l("MULTIL,0");                 // reset multilogs 
            WriteN4l("MULTIL,1,1,50");            // index 1 == phase1, rms voltage
            WriteN4l("MULTIL,2,2,50");            // index 2 == phase2, rms voltage
            WriteN4l("MULTIL,3,3,50");            // index 3 == phase3, rms voltage       
            WriteN4l("MULTIL,4,1,1");             // index 4 == phase1, frequency
            WriteN4l("MULTIL,5,1,79");            // index 5 == phase-phase, rms voltage 
            WriteN4l("REZERO");                   // request the dsp to re-compensate for dc offset
        }

        private void SpitPhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FREQ,60;:FREQ:SPAN,600");
            Thread.Sleep(8000);
            Write_PPS(":VOLT,0;:FORM,2;:VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SinglePhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,1;:FREQ," + UserFreq.Trim() +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SinglePhaseFreqRespControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,1;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
            CustomRest();
        }

        private void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT1?");
                var response = Read_Pps();
                if (double.TryParse(response, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 15;
        }

        private void PpsLoopTwoAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT2?");
                var response = Read_Pps();
                if (double.TryParse(response, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 15;
        }

        private void PpsLoopThreeAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT3?");
                var response = Read_Pps();
                if (double.TryParse(response, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 15;
        }

        private void PpsLoopAcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VLL1?");
                var response = Read_Pps();
                if (double.TryParse(response, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 15;
        }

        private void N4lMeasOne()
        {
            meas1 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                var response = Read_N4l().Split(',');
                if (response.Length > 0 && double.TryParse(response[0], out double value))
                {
                    PhaseA = response[0];
                    meas1 += value;
                }
            }
            meas1 /= 30;
        }

        private void N4lMeasTwo()
        {
            meas2 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                var response = Read_N4l().Split(',');
                if (response.Length > 1 && double.TryParse(response[1], out double value))
                {
                    PhaseB = response[1];
                    meas2 += value;
                }
            }
            meas2 /= 30;
        }

        private void N4lMeasThree()
        {
            meas3 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                var response = Read_N4l().Split(',');
                if (response.Length > 2 && double.TryParse(response[2], out double value))
                {
                    PhaseC = response[2];
                    meas3 += value;
                }
            }
            meas3 /= 30;
        }

        private void N4lMeasFour()
        {
            meas4 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                var response = Read_N4l().Split(',');
                if (response.Length > 3 && double.TryParse(response[3], out double value))
                {
                    Frequency = response[3];
                    meas4 += value;
                }
            }
            meas4 /= 30;
        }

        private void N4lMeasLineLine()
        {
            measlL = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                var response = Read_N4l().Split(',');
                if (response.Length > 4 && double.TryParse(response[4], out double value))
                {
                    PhaseAb = response[4];
                    measlL += value;
                }
            }
            measlL /= 30;
        }

        private double GetExcelCellValue(Excel.Range cell)
        {
            if (cell?.Value2 != null)
            {
                return Convert.ToDouble(cell.Value2);
            }
            return 0.0;
        }

        public void RunTest()
        {
            // opens workbook master and test worksheets          
            string mysheet = @"\\guimain\ATE\CalTestGUI\MasterPpsTest.xlsx";
            Excel.Application? xlApp = null;
            Excel.Workbook? sheet = null;

            try
            {
                xlApp = new Excel.Application();
                Excel.Workbooks books = xlApp.Workbooks;
                sheet = books.Open(mysheet);
                xlApp.DisplayFullScreen = true;
                xlApp.Visible = true;

                Excel.Worksheet? TwoPh = xlApp.Worksheets["2 Ph"] as Excel.Worksheet;
                Excel.Worksheet? FrontPage = xlApp.Worksheets["Front Page"] as Excel.Worksheet;

                if (TwoPh == null || FrontPage == null)
                {
                    MessageBox.Show("Required worksheets not found");
                    return;
                }

                // front page information and start test
                FrontPage.Activate();
                if (TwoPh.Cells[25, 5] is Excel.Range freqCell)
                {
                    freqCell.Value2 = TestFreq;
                }

                date = DateTime.UtcNow.ToLongDateString();
                time = DateTime.UtcNow.ToLongTimeString();

                if (FrontPage.Cells[11, 8] is Excel.Range timeCell)
                {
                    timeCell.Value2 = time;
                }
                if (FrontPage.Cells[17, 3] is Excel.Range dateCell)
                {
                    dateCell.Value2 = date;
                }
                if (FrontPage.Cells[19, 3] is Excel.Range modelCell)
                {
                    modelCell.Value2 = "118ACX-UPC1";
                }
                if (FrontPage.Cells[18, 3] is Excel.Range companyCell)
                {
                    companyCell.Value2 = Company;
                }
                if (FrontPage.Cells[20, 3] is Excel.Range serNumCell)
                {
                    serNumCell.Value2 = serNum;
                }
                if (FrontPage.Cells[21, 3] is Excel.Range caseNumCell)
                {
                    caseNumCell.Value2 = CaseNum;
                }
                if (FrontPage.Cells[17, 7] is Excel.Range certNumCell)
                {
                    certNumCell.Value2 = CertNum;
                }
                if (FrontPage.Cells[18, 7] is Excel.Range plantNumCell)
                {
                    plantNumCell.Value2 = PlantNum;
                }

                //split phase ac voltage test 
                TwoPh.Activate();
                Newtons4thControls();
                SpitPhaseAcControls();

                // Process cells 78-85
                for (int row = 78; row <= 85; row++)
                {
                    if (TwoPh.Cells[row, 2] is Excel.Range inputCell)
                    {
                        cellvalue = GetExcelCellValue(inputCell);
                        Write_PPS(":VOLT," + cellvalue);
                        CustomRest();
                        PpsLoopAcLineLine();
                        
                        if (TwoPh.Cells[row, 4] is Excel.Range outputCell1)
                        {
                            outputCell1.Value2 = pps1.ToString();
                        }
                        
                        N4lMeasLineLine();
                        if (TwoPh.Cells[row, 8] is Excel.Range outputCell2)
                        {
                            outputCell2.Value2 = measlL.ToString();
                        }
                    }
                }

                //single phase ac voltage test
                SinglePhaseAcControls();

                // Process cells 128-135
                for (int row = 128; row <= 135; row++)
                {
                    if (TwoPh.Cells[row, 2] is Excel.Range inputCell)
                    {
                        cellvalue = GetExcelCellValue(inputCell);
                        Write_PPS(":VOLT," + cellvalue);
                        CustomRest();
                        PpsLoopOneAcTest();
                        
                        if (TwoPh.Cells[row, 4] is Excel.Range outputCell1)
                        {
                            outputCell1.Value2 = pps1.ToString();
                        }
                        
                        N4lMeasOne();
                        if (TwoPh.Cells[row, 8] is Excel.Range outputCell2)
                        {
                            outputCell2.Value2 = PhaseA;
                        }
                    }
                }

                //frequency response
                SinglePhaseFreqRespControls();

                // Process cells 177-181
                for (int row = 177; row <= 181; row++)
                {
                    if (TwoPh.Cells[row, 2] is Excel.Range inputCell)
                    {
                        cellvalue = GetExcelCellValue(inputCell);
                        Write_PPS(":FREQ," + cellvalue);
                        CustomRest();
                        N4lMeasFour();
                        
                        if (TwoPh.Cells[row, 4] is Excel.Range outputCell1)
                        {
                            outputCell1.Value2 = meas4.ToString();
                        }
                        
                        N4lMeasOne();
                        if (TwoPh.Cells[row, 8] is Excel.Range outputCell2)
                        {
                            outputCell2.Value2 = meas1.ToString();
                        }
                    }
                }

                Write_PPS(":OUTP,OFF");
                CustomRest();
                Write_PPS(":FREQ:SPAN,1200");
                Thread.Sleep(8000);
                Write_PPS(":FREQ,1000");
                Write_PPS(":OUTP,ON");
                CustomRest();

                // Process cells 180-181 again with new settings
                for (int row = 180; row <= 181; row++)
                {
                    if (TwoPh.Cells[row, 2] is Excel.Range inputCell)
                    {
                        cellvalue = GetExcelCellValue(inputCell);
                        Write_PPS(":FREQ," + cellvalue);
                        CustomRest();
                        N4lMeasFour();
                        
                        if (TwoPh.Cells[row, 4] is Excel.Range outputCell1)
                        {
                            outputCell1.Value2 = meas4.ToString();
                        }
                        
                        N4lMeasOne();
                        if (TwoPh.Cells[row, 8] is Excel.Range outputCell2)
                        {
                            outputCell2.Value2 = meas1.ToString();
                        }
                    }
                }

                Write_PPS(":OUTP,OFF");
                CustomRest();
                Write_PPS(":FREQ,60");
                CustomRest();
                Write_PPS(":FREQ:SPAN,600");
                Thread.Sleep(8000);

                //reading k-factors
                Write_PPS(":CAL:KFACTORS:ALL?");
                string K_Factors = Read_Pps();

                //End off test
                FrontPage.Activate();
                time = DateTime.UtcNow.ToLongTimeString();
                if (FrontPage.Cells[12, 8] is Excel.Range endTimeCell)
                {
                    endTimeCell.Value2 = time;
                }

                Write_PPS(":OUTP,OFF");
                WriteN4l("KEYBOARD,ENABLE");
                Write_PPS("*GTL");
                Write_PPS("*RST");
                WriteN4l("*RST");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during test: {ex.Message}");
            }
            finally
            {
                if (sheet != null)
                {
                    sheet.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
            }
        }
    }
}
