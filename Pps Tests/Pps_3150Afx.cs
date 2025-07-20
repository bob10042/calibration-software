using NationalInstruments.Visa;
using System;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI
{
    /*
     External class to test the Pacific 3150AFX. 
     All communications and external equipment is set
     up through this case. 
    */
    public class Pps_3150Afx
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
            WriteN4l("MULTIL,6,1,58");            // index 6 == phase1, dc voltage
            WriteN4l("MULTIL,7,2,58");            // index 7 == phase2, dc voltage
            WriteN4l("MULTIL,8,3,58");            // index 8 == phase3, dc voltage
        }

        private void ThreePhaseControlsAC()
        {
            Write_PPS("OUTPUT:AUTO,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT:MODE,AC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:FREQ,400;:VOLT:AC:LIM:MIN,0;" +
                ":VOLT:AC:LIM:MAX,300;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");       
        }

        private void ThreePhaseFreqRespControls()
        {
            Write_PPS("OUTPUT:AUTO,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT:MODE,AC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:FREQ,400;:VOLT:AC:LIM:MIN,0;" +
                ":VOLT:AC:LIM:MAX,150;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");                
        }

        private void ThreePhaseControlsDC()
        {
            Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:DC:LIM:MIN,0;:VOLT:DC:LIM:MAX,425;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");            
        }

        private void SpitPhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":PROG:EXEC,0;:VOLT:MODE,AC;:FORM,2;:FREQ,400;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPL,DIRECT;:PHAS2,180;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");           
        }

        private void SpitPhaseDcControls()
        {
            Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,2;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":PHAS2,180;:RANG,1;:RAMP,0;:OUTP,ON");           
        }

        private void SinglePhaseAcControls()
        {
            Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":PROG:EXEC,0;:VOLT:MODE,AC;:FORM,1;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");
        }
        private void SinglePhaseDcControls()
        {
            Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,1;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");
        }

        private void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT:AC1?");
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
                Write_PPS(":MEAS:VOLT:AC2?");
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
                Write_PPS(":MEAS:VOLT:AC3?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopOneDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT:DC1?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopTwoDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT:DC2?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopThreeDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT:DC3?");
                pps1 += Convert.ToDouble(Read_Pps());
            }
            pps1 /= 10;
        }

        private void PpsLoopDcLineLine()
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

        private void N4lMeasOneAc()
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

        private void N4lMeasTwoAc()
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

        private void N4lMeasThreeAc()
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
        private void N4lMeasOneDc()
        {
            meas1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseA = Read_N4l().Split(',')[5];
                meas1 += Convert.ToDouble(PhaseA);
            }
            meas1 /= 10;
        }

        private void N4lMeasTwoDc()
        {
            meas2 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseB = Read_N4l().Split(',')[6];
                meas2 += Convert.ToDouble(PhaseB);
            }
            meas2 /= 10;
        }

        private void N4lMeasThreeDc()
        {
            meas3 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseC = Read_N4l().Split(',')[7];
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
            Excel.Worksheet ThreePh = xlApp.Worksheets["3 Ph AFX"] as Microsoft.Office.Interop.Excel.Worksheet;
            Excel.Worksheet FrontPage = xlApp.Worksheets["Front Page"] as Microsoft.Office.Interop.Excel.Worksheet;

            // front page information and start test
            FrontPage.Activate();
            ThreePh.Cells[25, 5] = TestFreq;
            date = DateTime.UtcNow.ToLongDateString();
            time = DateTime.UtcNow.ToLongTimeString();
            FrontPage.Cells[11, 8] = time;
            FrontPage.Cells[17, 3] = date;
            FrontPage.Cells[19, 3] = "3150AFX";
            FrontPage.Cells[18, 3] = Company;
            FrontPage.Cells[20, 3] = serNum;
            FrontPage.Cells[21, 3] = CaseNum;
            FrontPage.Cells[17, 7] = CertNum;
            FrontPage.Cells[18, 7] = PlantNum;

            //three phase Ac voltage test 
            ThreePh.Activate();
            Newtons4thControls();
            ThreePhaseControlsAC();
            cellvalue = (ThreePh.Cells[77, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);           
            Thread.Sleep(30000);
            PpsLoopOneAcTest();
            ThreePh.Cells[77, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[77, 8] = Convert.ToString(meas1);            
            PpsLoopTwoAcTest();
            ThreePh.Cells[78, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[78, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[79, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[79, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[80, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);                     
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[80, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[80, 8] = Convert.ToString(meas1);            
            PpsLoopTwoAcTest();
            ThreePh.Cells[81, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[81, 8] = Convert.ToString(meas2);            
            PpsLoopThreeAcTest();
            ThreePh.Cells[82, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[82, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[83, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[83, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[83, 8] = Convert.ToString(meas1);       
            PpsLoopTwoAcTest();
            ThreePh.Cells[84, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[84, 8] = Convert.ToString(meas2);       
            PpsLoopThreeAcTest();
            ThreePh.Cells[85, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[85, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[86, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[86, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[86, 8] = Convert.ToString(meas1);         
            PpsLoopTwoAcTest();
            ThreePh.Cells[87, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[87, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[88, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[88, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[89, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[89, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[89, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[90, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[90, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[91, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[91, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[92, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);           
            PpsLoopOneAcTest();
            ThreePh.Cells[92, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[92, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[93, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[93, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[94, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[94, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[95, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[95, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[95, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[96, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[96, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[97, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[97, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[98, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            PpsLoopOneAcTest();
            ThreePh.Cells[98, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[98, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[99, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[99, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[100, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[100, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[118, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[118, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[118, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[119, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[119, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[120, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[120, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[121, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[121, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[121, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[122, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[122, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[123, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[123, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[124, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[124, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[124, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[125, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[125, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[126, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[126, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[127, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneAcTest();
            ThreePh.Cells[127, 4] = Convert.ToString(pps1);
            N4lMeasOneAc();
            ThreePh.Cells[127, 8] = Convert.ToString(meas1);
            PpsLoopTwoAcTest();
            ThreePh.Cells[128, 4] = Convert.ToString(pps1);
            N4lMeasTwoAc();
            ThreePh.Cells[128, 8] = Convert.ToString(meas2);
            PpsLoopThreeAcTest();
            ThreePh.Cells[129, 4] = Convert.ToString(pps1);
            N4lMeasThreeAc();
            ThreePh.Cells[129, 8] = Convert.ToString(meas3);



            ////// add disable the voltage range for new firmware. otherwise the frequency doesn't go above 800Hz



            //frequency response
            ThreePhaseFreqRespControls();
            Write_PPS(":VOLT:AC,115");
            cellvalue = (ThreePh.Cells[217, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(30000);
            N4lMeasFour();
            ThreePh.Cells[217, 4] = Convert.ToString(meas4);
            N4lMeasOneAc();
            ThreePh.Cells[217, 8] = Convert.ToString(meas1);
            N4lMeasTwoAc();
            ThreePh.Cells[218, 8] = Convert.ToString(meas2);
            N4lMeasThreeAc();
            ThreePh.Cells[219, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[220, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[220, 4] = Convert.ToString(meas4);
            N4lMeasOneAc();
            ThreePh.Cells[220, 8] = Convert.ToString(meas1);
            N4lMeasTwoAc();
            ThreePh.Cells[221, 8] = Convert.ToString(meas2);
            N4lMeasThreeAc();
            ThreePh.Cells[222, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[223, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[223, 4] = Convert.ToString(meas4);
            N4lMeasOneAc();
            ThreePh.Cells[223, 8] = Convert.ToString(meas1);
            N4lMeasTwoAc();
            ThreePh.Cells[224, 8] = Convert.ToString(meas2);
            N4lMeasThreeAc();
            ThreePh.Cells[225, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[226, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[226, 4] = Convert.ToString(meas4);
            N4lMeasOneAc();
            ThreePh.Cells[226, 8] = Convert.ToString(meas1);
            N4lMeasTwoAc();
            ThreePh.Cells[227, 8] = Convert.ToString(meas2);
            N4lMeasThreeAc();
            ThreePh.Cells[228, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[229, 2] as Excel.Range).Value;
            Write_PPS(":FREQ," + cellvalue);
            Thread.Sleep(15000);
            N4lMeasFour();
            ThreePh.Cells[229, 4] = Convert.ToString(meas4);
            N4lMeasOneAc();
            ThreePh.Cells[229, 8] = Convert.ToString(meas1);
            N4lMeasTwoAc();
            ThreePh.Cells[230, 8] = Convert.ToString(meas2);
            N4lMeasThreeAc();
            ThreePh.Cells[231, 8] = Convert.ToString(meas3);

            //split phase ac voltage test
            SpitPhaseAcControls();
            cellvalue = (ThreePh.Cells[144, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:AC," + cellvalue);
            Thread.Sleep(30000);
            PpsLoopAcLineLine();
            ThreePh.Cells[144, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[144, 8] = Convert.ToString(measlL);

            for (int i = 144; i < 152; i++)
            {
                cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                Write_PPS(":VOLT:AC," + cellvalue);
                Thread.Sleep(15000);
                PpsLoopAcLineLine();
                ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                N4lMeasLineLine();
                ThreePh.Cells[i, 8] = Convert.ToString(measlL);
            }

            //three phase dc voltage test
            ThreePhaseControlsDC();
            ThreePh.Activate();
            Newtons4thControls();
            cellvalue = (ThreePh.Cells[363, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(30000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[363, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[363, 8] = Convert.ToString(meas1);           
            PpsLoopTwoDcTest();
            ThreePh.Cells[364, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[364, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[365, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[365, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[366, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[366, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[366, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[367, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[367, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[368, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[368, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[369, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[369, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[369, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[370, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[370, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[371, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[371, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[372, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[372, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[372, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[373, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[373, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[374, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[374, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[375, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[375, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[375, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[376, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[376, 8] = Convert.ToString(meas2);            
            PpsLoopThreeDcTest();
            ThreePh.Cells[377, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[377, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[378, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);            
            PpsLoopOneDcTest();
            ThreePh.Cells[378, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[378, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[379, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[379, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[380, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[380, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[381, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneDcTest();
            ThreePh.Cells[381, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[381, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[382, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[382, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[383, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[383, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[384, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneDcTest();
            ThreePh.Cells[384, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[384, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[385, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[385, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[386, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[386, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[404, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneDcTest();
            ThreePh.Cells[404, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[404, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[405, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[405, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[406, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[406, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[407, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);
            PpsLoopOneDcTest();
            ThreePh.Cells[407, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[407, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[408, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[408, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[409, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[409, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[410, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000); //changed for error checking
            PpsLoopOneDcTest();
            ThreePh.Cells[410, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[410, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[411, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[411, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[412, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[412, 8] = Convert.ToString(meas3);

            cellvalue = (ThreePh.Cells[413, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(15000);  //changed for error checking
            PpsLoopOneDcTest();
            ThreePh.Cells[413, 4] = Convert.ToString(pps1);
            N4lMeasOneDc();
            ThreePh.Cells[413, 8] = Convert.ToString(meas1);
            PpsLoopTwoDcTest();
            ThreePh.Cells[414, 4] = Convert.ToString(pps1);
            N4lMeasTwoDc();
            ThreePh.Cells[414, 8] = Convert.ToString(meas2);
            PpsLoopThreeDcTest();
            ThreePh.Cells[415, 4] = Convert.ToString(pps1);
            N4lMeasThreeDc();
            ThreePh.Cells[415, 8] = Convert.ToString(meas3);

            //split phase dc voltage test 
            SpitPhaseDcControls();
            cellvalue = (ThreePh.Cells[430, 2] as Excel.Range).Value;
            Write_PPS(":VOLT:DC," + cellvalue);
            Thread.Sleep(30000);
            PpsLoopDcLineLine();
            ThreePh.Cells[430, 4] = Convert.ToString(pps1);
            N4lMeasLineLine();
            ThreePh.Cells[430, 8] = Convert.ToString(measlL);

            for (int i = 430; i < 439; i++)
            {
                cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                Write_PPS(":VOLT:DC," + cellvalue);
                Thread.Sleep(15000);
                PpsLoopDcLineLine();
                ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                N4lMeasLineLine();
                ThreePh.Cells[i, 8] = Convert.ToString(measlL);
            }         
            
            //single phase AC
            Write_PPS(":OUTP,OFF");
            string message = "Link out the three phase output, Cancel to stop test";
            string title = "Single Phase Test";
            MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.OK)
            {
                //single phase AC test
                SinglePhaseAcControls();
                ThreePh.Activate();
                Newtons4thControls();
                cellvalue = (ThreePh.Cells[171, 2] as Excel.Range).Value;
                Write_PPS(":VOLT:AC," + cellvalue);
                Thread.Sleep(30000);
                PpsLoopOneAcTest();
                ThreePh.Cells[171, 4] = Convert.ToString(pps1);
                N4lMeasOneAc();
                ThreePh.Cells[171, 8] = Convert.ToString(meas1);

                for (int i = 171; i < 183; i++)
                {
                    cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                    Write_PPS(":VOLT:AC," + cellvalue);
                    Thread.Sleep(15000);
                    PpsLoopOneAcTest();
                    ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                    N4lMeasOneAc();
                    ThreePh.Cells[i, 8] = Convert.ToString(meas1);
                }

                //single phase DC test
                SinglePhaseDcControls();
                ThreePh.Activate();
                cellvalue = (ThreePh.Cells[456, 2] as Excel.Range).Value;
                Write_PPS(":VOLT:DC," + cellvalue);
                Thread.Sleep(30000);
                PpsLoopOneDcTest();
                ThreePh.Cells[456, 4] = Convert.ToString(pps1);
                N4lMeasOneDc();
                ThreePh.Cells[456, 8] = Convert.ToString(meas1);

                for (int i = 456; i < 466; i++)
                {
                    cellvalue = (ThreePh.Cells[i, 2] as Excel.Range).Value;
                    Write_PPS(":VOLT:DC," + cellvalue);
                    Thread.Sleep(15000);
                    PpsLoopOneDcTest();
                    ThreePh.Cells[i, 4] = Convert.ToString(pps1);
                    N4lMeasOneDc();
                    ThreePh.Cells[i, 8] = Convert.ToString(meas1);
                }

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
            else
            {
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
}
