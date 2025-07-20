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
     External class to test the Pacific 360AMX_UPC32. 
     All communications and external equipment is set
     up through this case. 
    */
    public class Pps_360AmxUpc_32
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

        private void ThreePhaseControls()
        {                                       
            Write_PPS("*LLO");                            
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,3;:FREQ," + UserFreq.Trim() +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SpitPhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,2;:FREQ,50;" +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void SinglePhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,1;:FREQ," + UserFreq.Trim() +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            CustomRest();
        }

        private void ThreePhaseFreqRespControls()
        {
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,3;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHASe2,120;" +
                ":PHASe3,240;:RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
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
            Excel.Worksheet ThreePh = xlApp.Worksheets["3 Ph"] as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Microsoft.Office.Interop.Excel.Worksheet;

            // front page information and start test
            FrontPage.Activate();
            ThreePh.Cells[25, 5] = TestFreq;
            date = DateTime.UtcNow.ToLongDateString();
            time = DateTime.UtcNow.ToLongTimeString();
            FrontPage.Cells[11, 8] = time;
            FrontPage.Cells[17, 3] = date;
            FrontPage.Cells[19, 3] = "360AMX-UPC32";
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
            Write_PPS(":VOLT,"+cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[77, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[77, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[78, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[78, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[78, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[79, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[79, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[79, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[80, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[80, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[80, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[81, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[81, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[81, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[82, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[82, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[82, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[83, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[83, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[83, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[84, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[84, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[84, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[85, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[85, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[85, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[86, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[86, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[86, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[87, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[87, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[87, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[88, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[88, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[88, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[89, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[89, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[89, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[90, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[90, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[90, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[91, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[91, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[91, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[92, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[92, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[92, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[93, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[93, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[93, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[94, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[94, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[94, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[95, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[95, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[95, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[96, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[96, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[96, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[97, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[97, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[97, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[98, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[98, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[98, 8] = Convert.ToString(meas1);
            cellvalue = (ThreePh.Cells[99, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopTwoAcTest();
            ThreePh.Cells[99, 4] = Convert.ToString(pps1);
            N4lMeasTwo();
            ThreePh.Cells[99, 8] = Convert.ToString(meas2);
            cellvalue = (ThreePh.Cells[100, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopThreeAcTest();
            ThreePh.Cells[100, 4] = Convert.ToString(pps1);
            N4lMeasThree();
            ThreePh.Cells[100, 8] = Convert.ToString(meas3);

            //split phase ac voltage test
            SpitPhaseAcControls();
            cellvalue = (ThreePh.Cells[126, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[126, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[126, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[127, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[127, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[127, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[128, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[128, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[128, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[129, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[129, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[129, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[130, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[130, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[130, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[131, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[131, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[131, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[132, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[132, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[132, 8] = Convert.ToString(measlL);

            cellvalue = (ThreePh.Cells[133, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopAcLineLine();
            ThreePh.Cells[133, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[133, 8] = Convert.ToString(measlL);

            //single phase ac voltage test
            SinglePhaseAcControls();
            cellvalue = (ThreePh.Cells[176, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[176, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[176, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[177, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[177, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[177, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[178, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[178, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[178, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[179, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[179, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[179, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[180, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[180, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[180, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[181, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[181, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[181, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[182, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[182, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[182, 8] = PhaseA;

            cellvalue = (ThreePh.Cells[183, 2] as Excel.Range).Value;
            Write_PPS(":VOLT," + cellvalue);
            CustomRest();
            PpsLoopOneAcTest();
            ThreePh.Cells[183, 4] = Convert.ToString(pps1);
            N4lMeasOne();
            ThreePh.Cells[183, 8] = PhaseA;

            //frequency response
            ThreePhaseFreqRespControls();
            cellvalue = (ThreePh.Cells[226, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
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
            CustomRest();
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
            CustomRest();
            N4lMeasFour();
            ThreePh.Cells[232, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[232, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[233, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[234, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[235, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            CustomRest();
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
            CustomRest();
            N4lMeasFour();
            ThreePh.Cells[238, 4] = Convert.ToString(meas4);
            N4lMeasOne();
            ThreePh.Cells[238, 8] = Convert.ToString(meas1);
            N4lMeasTwo();
            ThreePh.Cells[239, 8] = Convert.ToString(meas2);
            N4lMeasThree();
            ThreePh.Cells[240, 8] = Convert.ToString(meas3);

            //reading k-factors
            Write_PPS(":CAL:KFACTORS:ALL?");
            string K_Factors = Read_Pps();
            string KinA = K_Factors.Split(',')[0];            
            string KinB = K_Factors.Split(',')[1];            
            string KinC = K_Factors.Split(',')[2];            
            string KexA = K_Factors.Split(',')[3];            
            string KexB = K_Factors.Split(',')[4];            
            string KexC = K_Factors.Split(',')[5];            
            string KiA = K_Factors.Split(',')[6];            
            string KiB = K_Factors.Split(',')[7];            
            string KiC = K_Factors.Split(',')[8];            
            string K1P = K_Factors.Split(',')[9];          
            ThreePh.Cells[317, 3] = KinA;
            ThreePh.Cells[317, 6] = KinB;
            ThreePh.Cells[317, 9] = KinC;
            ThreePh.Cells[318, 3] = KexA;
            ThreePh.Cells[318, 6] = KexB;
            ThreePh.Cells[318, 9] = KexC;
            ThreePh.Cells[319, 3] = KiA;
            ThreePh.Cells[319, 6] = KiB;
            ThreePh.Cells[319, 9] = KiC;
            ThreePh.Cells[317, 10] = K1P;

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
