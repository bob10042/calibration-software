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
     External class to test the Pacific 115ACX_UPC1. 
     All communications and external equipment is set
     up through this case. 
    */
    class Pps_115AcxUpc_1
    {
        private SerialPort SerWriteN4l = new SerialPort();
        public MessageBasedSession mbSession = null;
        public string visaName;
        public string TestFreq;
        public string Company;
        public string CertNum;
        public string CaseNum;
        public string serNum;
        public string UserFreq;
        public string PlantNum;

        string date = null;
        string time = null;
        string PhaseA = null;
        string PhaseB = null;
        string PhaseC = null;
        string Frequency = null;
        string PhaseAb = null;

        double meas1 = 0;
        double meas2 = 0;
        double meas3 = 0;
        double meas4 = 0;
        double measlL = 0;
        double pps1 = 0;
        double cellvalue = 0;

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
            if (visaName == null)
            {
                MessageBox.Show("Select Visa Address");
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
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        private string Read_Pps()
        {
            try
            {
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
            else
            { }
        }

        public string PortName_Newtowns_4th
        {
            get
            {
                return SerWriteN4l.PortName;
            }
            set
            {
                SerWriteN4l.PortName = value;
            }
        }

        public bool IsOpen_Newtowns_4th
        {
            get
            {
                return SerWriteN4l.IsOpen;
            }
        }

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

        public string Read_N4l()
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
            //3Write_PPS("*LLO");
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
                pps1 += Convert.ToDouble(Read_Pps());
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
                pps1 += Convert.ToDouble(Read_Pps());
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
                pps1 += Convert.ToDouble(Read_Pps());
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
                pps1 += Convert.ToDouble(Read_Pps());
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
                PhaseA = Read_N4l().Split(',')[0];
                meas1 += Convert.ToDouble(PhaseA);
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
                PhaseB = Read_N4l().Split(',')[1];
                meas2 += Convert.ToDouble(PhaseB);
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
                PhaseC = Read_N4l().Split(',')[2];
                meas3 += Convert.ToDouble(PhaseC);
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
                Frequency = Read_N4l().Split(',')[3];
                meas4 += Convert.ToDouble(Frequency);
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
                PhaseAb = Read_N4l().Split(',')[4];
                measlL += Convert.ToDouble(PhaseAb);
            }
            measlL /= 30;
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
            Excel.Worksheet TwoPh = xlApp.Worksheets["2 Ph"] as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Microsoft.Office.Interop.Excel.Worksheet;

            // front page information and start test
            FrontPage.Activate();
            TwoPh.Cells[25, 5] = TestFreq;
            date = DateTime.UtcNow.ToLongDateString();
            time = DateTime.UtcNow.ToLongTimeString();
            FrontPage.Cells[11, 8] = time;
            FrontPage.Cells[17, 3] = date;
            FrontPage.Cells[19, 3] = "115ACX-UPC1";
            FrontPage.Cells[18, 3] = Company;
            FrontPage.Cells[20, 3] = serNum;
            FrontPage.Cells[21, 3] = CaseNum;
            FrontPage.Cells[17, 7] = CertNum;
            FrontPage.Cells[18, 7] = PlantNum;

            //split phase ac voltage test 
            TwoPh.Activate();
            Newtons4thControls();
            SpitPhaseAcControls();
            cellvalue = (TwoPh.Cells[78, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[78, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[78, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[79, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[79, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[79, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[80, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[80, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[80, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[81, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[81, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[81, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[82, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[82, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[82, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[83, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[83, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[83, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[84, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[84, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[84, 8] = Convert.ToString(measlL);

            cellvalue = (TwoPh.Cells[85, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            TwoPh.Cells[85, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            TwoPh.Cells[85, 8] = Convert.ToString(measlL);

            //single phase ac voltage test
            SinglePhaseAcControls();
            cellvalue = (TwoPh.Cells[128, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[128, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[128, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[129, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[129, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[129, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[130, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[130, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[130, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[131, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[131, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[131, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[132, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[132, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[132, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[133, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[133, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[133, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[134, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[134, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[134, 8] = PhaseA;

            cellvalue = (TwoPh.Cells[135, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            TwoPh.Cells[135, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            TwoPh.Cells[135, 8] = PhaseA;

            //frequency response
            SinglePhaseFreqRespControls();
            cellvalue = (TwoPh.Cells[177, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
            N4lMeasFour();
            TwoPh.Cells[177, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            TwoPh.Cells[177, 8] = Convert.ToString(meas1);

            cellvalue = (TwoPh.Cells[178, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
            N4lMeasFour();
            TwoPh.Cells[178, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            TwoPh.Cells[178, 8] = Convert.ToString(meas1);

            cellvalue = (TwoPh.Cells[179, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
            N4lMeasFour();
            TwoPh.Cells[179, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            TwoPh.Cells[179, 8] = Convert.ToString(meas1);

            Write_PPS(":OUTP,OFF");
            CustomRest();
            Write_PPS(":FREQ:SPAN,1200");
            Thread.Sleep(8000);
            Write_PPS(":OUTP,ON");
            CustomRest();

            cellvalue = (TwoPh.Cells[180, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
            N4lMeasFour();
            TwoPh.Cells[180, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            TwoPh.Cells[180, 8] = Convert.ToString(meas1);

            cellvalue = (TwoPh.Cells[181, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue); ;           
            CustomRest();
            N4lMeasFour();
            TwoPh.Cells[181, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            TwoPh.Cells[181, 8] = Convert.ToString(meas1);

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
            FrontPage.Cells[12, 8] = time;
            Write_PPS(":OUTP,OFF");
            WriteN4l("KEYBOARD,ENABLE");
            Write_PPS("*GTL");
            Write_PPS("*RST");
            WriteN4l("*RST");
        }
    }
}

