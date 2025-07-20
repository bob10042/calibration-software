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
     External class to test the Pacific 360ASX_UPC3. 
     All communications and external equipment is set
     up through this case. 
    */
    public class Pps_360AsxUpc_3
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
        }

        private void ThreePhaseControls()
        {
            Write_PPS("*LLO");
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FREQ,60;:FREQ:SPAN,600");
            Thread.Sleep(8000);
            Write_PPS(":VOLT,0;:FORM,3;:FREQ," + UserFreq.Trim() +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SpitPhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,2;:FREQ,50;" +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SinglePhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,1;:FREQ," + UserFreq.Trim() +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void ThreePhaseFreqRespControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,3;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1200;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHASe2,120;" +
                ":PHASe3,240;:RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
            CustomRest();
        }

        private void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT1?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopTwoAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT2?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopThreeAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT3?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopAcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VLL1?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void N4lMeasOne()
        {
            meas1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseA = Read_N4l().Split(',')[0];
                meas1 += Convert.ToDouble(PhaseA);
            }
            meas1 /= 10;
        }

        private void N4lMeasTwo()
        {
            meas2 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseB = Read_N4l().Split(',')[1];
                meas2 += Convert.ToDouble(PhaseB);
            }
            meas2 /= 10;
        }

        private void N4lMeasThree()
        {
            meas3 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseC = Read_N4l().Split(',')[2];
                meas3 += Convert.ToDouble(PhaseC);
            }
            meas3 /= 10;
        }

        private void N4lMeasFour()
        {
            meas4 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                Frequency = Read_N4l().Split(',')[3];
                meas4 += Convert.ToDouble(Frequency);
            }
            meas4 /= 10;
        }

        private void N4lMeasLineLine()
        {
            measlL = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseAb = Read_N4l().Split(',')[4];
                measlL += Convert.ToDouble(PhaseAb);
            }
            measlL /= 10;
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
            Excel.Worksheet ThreePh = xlApp.Worksheets["3 Ph"] as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Microsoft.Office.Interop.Excel.Worksheet;

            // front page information and start test
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

            //three phase ac voltage test 
            ThreePh.Activate();
            Newtons4thControls();
            ThreePhaseControls();
            cellvalue = (ThreePh.Cells[77, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(30000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[77, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[77, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[78, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[78, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[79, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[79, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[80, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[80, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[80, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[81, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[81, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[82, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[82, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[83, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[83, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[83, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[84, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[84, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[85, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[85, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[86, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[86, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[86, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[87, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[87, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[88, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[88, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[89, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[89, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[89, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[90, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[90, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[91, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[91, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[92, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[92, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[92, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[93, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[93, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[94, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[94, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[95, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[95, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[95, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[96, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[96, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[97, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[97, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[98, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneAcTest();
            ThreePh.Cells[98, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[98, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[99, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[99, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[100, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[100, 8] = Convert.ToString(meas3);

            //split phase ac voltage test
            SpitPhaseAcControls();
            cellvalue = (ThreePh.Cells[126, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(30000);
            PpsLoopAcLineLine();
            ThreePh.Cells[126, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[126, 8] = Convert.ToString(measlL);

            for (int i = 126; i < 134; i++)
            {
                cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                Write_PPS(":VOLT," + cellvalue);
                Thread.Sleep(15000);
                PpsLoopAcLineLine();
                ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                N4lMeasLineLine();
                ThreePh.Cells[i, 8] = Convert.ToString(measlL);
            }

            //single phase ac voltage test
            SinglePhaseAcControls();
            cellvalue = (ThreePh.Cells[176, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            Thread.Sleep(30000);
            PpsLoopOneAcTest();
            ThreePh.Cells[176, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[176, 8] = PhaseA;

            for (int i = 176; i < 184; i++)
            {
                cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                Write_PPS(":VOLT," + cellvalue);
                Thread.Sleep(15000);
                PpsLoopOneAcTest();
                ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                N4lMeasOne();
                ThreePh.Cells[i, 8] = Convert.ToString(meas1);
            }

            //frequency response
            ThreePhaseFreqRespControls();
            cellvalue = (ThreePh.Cells[226, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(30000);
            N4lMeasFour();
            ThreePh.Cells[226, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[226, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[227, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[228, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[229, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[229, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[229, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[230, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[231, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[232, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[232, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[232, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[233, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[234, 8] = Convert.ToString(meas3);

            Write_PPS(":OUTP,OFF");
            CustomRest();
            Write_PPS(":FREQ:SPAN,1200");
            Thread.Sleep(8000);
            Write_PPS(":FREQ,1000");
            Write_PPS(":OUTP,ON");
            CustomRest();

            cellvalue = (ThreePh.Cells[235, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(30000);
            N4lMeasFour();
            ThreePh.Cells[235, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[235, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[236, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[237, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[238, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[238, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[238, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[239, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[240, 8] = Convert.ToString(meas3);

            Write_PPS(":OUTP,OFF");
            CustomRest();
            Write_PPS(":FREQ,60");
            CustomRest();
            Write_PPS(":FREQ:SPAN,600");
            Thread.Sleep(8000);           
            
            //reading k-factors
            Write_PPS(":CAL:KFACTORS:ALL?");
            string K_Factors = Read_Pps();
            string SingleKin0 = K_Factors.Split(',')[0];
            string SingleKin2 = K_Factors.Split(',')[3];
            string SingleKin4 = K_Factors.Split(',')[6];
            string SingleKo6 = K_Factors.Split(',')[9];          
            ThreePh.Cells[325, 8] = SingleKin0;
            ThreePh.Cells[326, 8] = SingleKin2;
            ThreePh.Cells[327, 8] = SingleKin4;
            ThreePh.Cells[328, 8] = SingleKo6;
            string SplitKin8 = K_Factors.Split(',')[12];
            string SplitKin10 = K_Factors.Split(',')[15];
            string SplitKin12 = K_Factors.Split(',')[18];
            string SplitKo14 = K_Factors.Split(',')[21];
            ThreePh.Cells[325, 3] = SplitKin8;
            ThreePh.Cells[326, 3] = SplitKin10;
            ThreePh.Cells[327, 3] = SplitKin12;
            ThreePh.Cells[328, 3] = SplitKo14;
            string ThreeKinA = K_Factors.Split(',')[24];
            string ThreeKinB = K_Factors.Split(',')[25];
            string ThreeKinC = K_Factors.Split(',')[26];            
            ThreePh.Cells[317, 3] = ThreeKinA;
            ThreePh.Cells[317, 6] = ThreeKinB;
            ThreePh.Cells[317, 9] = ThreeKinC;         
            string ThreeKexA = K_Factors.Split(',')[27];
            string ThreeKexB = K_Factors.Split(',')[28];
            string ThreeKexC = K_Factors.Split(',')[29];
            ThreePh.Cells[318, 3] = ThreeKexA;
            ThreePh.Cells[318, 6] = ThreeKexB;
            ThreePh.Cells[318, 9] = ThreeKexC;
            string ThreeKiA = K_Factors.Split(',')[30];
            string ThreeKiB = K_Factors.Split(',')[31];
            string ThreeKiC = K_Factors.Split(',')[32];            
            ThreePh.Cells[319, 3] = ThreeKiA;
            ThreePh.Cells[319, 6] = ThreeKiB;
            ThreePh.Cells[319, 9] = ThreeKiC;
            string ThreeKoA = K_Factors.Split(',')[33];
            string ThreeKoB = K_Factors.Split(',')[34];
            string ThreeKoC = K_Factors.Split(',')[35];
            ThreePh.Cells[320, 3] = ThreeKoA;
            ThreePh.Cells[320, 6] = ThreeKoB;
            ThreePh.Cells[320, 9] = ThreeKoC;

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

