using System;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;

namespace CalTestGUI
{
    /*
     External class to test the Portable appliance test. 
     All communications and external equipment is set
     up through this case. 
    */
    public class GWI
    {
        private SerialPort serGWI = new SerialPort();
        public string ResLimit;
        public string CurrRating;
        public string VoltRating;

        public string PortName
        {
            get
            {
                return serGWI.PortName;
            }
            set
            {
                serGWI.PortName = value;
            }
        }

        public bool IsOpen
        {
            get
            {
                return serGWI.IsOpen;
            }
        }

        public void Open()
        {
            //Serial port setup
            serGWI.BaudRate = 115200;
            serGWI.Parity = Parity.None;
            serGWI.StopBits = StopBits.One;
            serGWI.DataBits = 8;
            serGWI.Handshake = Handshake.None;
            serGWI.RtsEnable = true;
            serGWI.DtrEnable = true;
            serGWI.WriteTimeout = 1000;
            serGWI.ReadTimeout = 1000;
            serGWI.Open();

            //sets up the gwinstek system settings
            serGWI.WriteLine("SYST:LCD:CONT 5");
            serGWI.WriteLine("SYST:LCD:BRIG 2");
            serGWI.WriteLine("SYST:BUZZ:PSOUND ON");
            serGWI.WriteLine("SYST:BUZZ:FSOUND ON");
            serGWI.WriteLine("SYST:BUZZ:PTIM 0.5");
            serGWI.WriteLine("SYST:BUZZ:FTIM 0.5");
            
            //sets up the gwinstek utility settings 
            serGWI.WriteLine("MANU:UTIL:GROUNDMODE OFF");
            serGWI.WriteLine("MANU:UTIL:ARCM OFF");
            serGWI.WriteLine("MANU:UTIL:PASS OFF");
            serGWI.WriteLine("MANU:UTIL:FAIL STOP");
            serGWI.WriteLine("MANU:UTIL:MAXH OFF");
        }

        public void Close()
        {
            serGWI.Close();
        }

        public string Read()
        {
            try
            {
                string message = serGWI.ReadLine();
                return (message);
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }
        public void ClosePort()
        {
            try
            {
                if (serGWI.IsOpen)
                {
                    serGWI.Close();
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string ChangeHighResLimit()
        {
            //Changes the resistance limit based on the length of the cable 
            try
            {
                if (ResLimit == "1 METRES") { ResLimit = "MANU:GB:RHIS 111"; }
                if (ResLimit == "2 METRES") { ResLimit = "MANU:GB:RHIS 112"; }
                if (ResLimit == "3 METRES") { ResLimit = "MANU:GB:RHIS 112"; }
                if (ResLimit == "4 METRES") { ResLimit = "MANU:GB:RHIS 113"; }
                if (ResLimit == "5 METRES") { ResLimit = "MANU:GB:RHIS 114"; }
                if (ResLimit == "6 METRES") { ResLimit = "MANU:GB:RHIS 115"; }
                if (ResLimit == "7 METRES") { ResLimit = "MANU:GB:RHIS 116"; }
                if (ResLimit == "8 METRES") { ResLimit = "MANU:GB:RHIS 116"; }
                if (ResLimit == "9 METRES") { ResLimit = "MANU:GB:RHIS 117"; }
                if (ResLimit == "10 METRES") { ResLimit = "MANU:GB:RHIS 118"; }
                if (ResLimit == "45 METRES") { ResLimit = "MANU:GB:RHIS 460"; }
                else if (ResLimit == null) { ResLimit = "MANU:GB:RHIS 100"; }
                return ResLimit;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string ChangeCurrRating()
        {
            try
            {
                if (CurrRating == "3 AMPS") { CurrRating = "MANU:GB:CURR 03.00"; }
                if (CurrRating == "5 AMPS") { CurrRating = "MANU:GB:CURR 05.00"; }
                if (CurrRating == "10 AMPS") { CurrRating = "MANU:GB:CURR 10.00"; }
                if (CurrRating == "25 AMPS") { CurrRating = "MANU:GB:CURR 25.00"; }
                return CurrRating;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string ChangeVoltRating()
        {
            if (VoltRating == "250 VOLTS") { VoltRating = "MANU:IR:VOLT 0.25"; }
            if (VoltRating == "500 VOLTS") { VoltRating = "MANU:IR:VOLT 0.5"; }
            return VoltRating;
        }
        public void StopTest()
        {
            serGWI.WriteLine("*RST");
            serGWI.WriteLine("*CLS");
            serGWI.WriteLine("FUNC:TEST OFF");
            serGWI.WriteLine("*RMTOFF");
        }

        public void GroundBondTest()
        {
            serGWI.WriteLine("MAIN:FUNC MANU");
            serGWI.WriteLine("MANU:STEP 1");
            serGWI.WriteLine("MANU:NAME EarthBond");
            serGWI.WriteLine("MANU:EDIT:MODE GB");
            serGWI.WriteLine("MANU:GB:FREQ 50");
            serGWI.WriteLine(ChangeHighResLimit());
            serGWI.WriteLine("MANU:GB:RLOS 0");
            serGWI.WriteLine(ChangeCurrRating());
            serGWI.WriteLine("MANU:GB:TTIM 20");
            serGWI.WriteLine("MANU:GB:REF 11.3");
            serGWI.WriteLine("MANU:GB:ZEROCHECK ON");
            serGWI.WriteLine("MAIN:FUNC MANU");
            serGWI.WriteLine("FUNC:TEST ON");
            Thread.Sleep(21000);
        }

        public void InsulationResistanceTest()
        {
            serGWI.WriteLine("MAIN:FUNC MANU");
            serGWI.WriteLine("MANU:STEP 2");
            serGWI.WriteLine("MANU:NAME Insulation");
            serGWI.WriteLine("MANU:EDIT:MODE IR");
            serGWI.WriteLine(ChangeVoltRating());
            serGWI.WriteLine("MANU:IR:RHIS 1.00");
            serGWI.WriteLine("MANU:IR:RLOS 1");
            serGWI.WriteLine("MANU:IR:TTIM 20");
            serGWI.WriteLine("MANU:IR:REF 0");
            serGWI.WriteLine("MAIN:FUNC MANU");
            serGWI.WriteLine("FUNC:TEST ON");
            Thread.Sleep(25000);
        }

        public void Class_1()
        {
            //opens a excel speadsheet prewriten template
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"\\guimain\ATE\CalTestGUI\MasterPatTest.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            excel.Visible = true;
                        
            string visPass = "PASS";
            //string MbsPas = "PASS";
            string ClassOne = "Class 1";
            x.Cells[36, 9] = visPass;
            //x.Cells[37, 9] = MbsPas;
            GroundBondTest();
            //writes data the the exel speadsheet    
            serGWI.WriteLine("MEAS?");
            string myStringEGB = Read();
            string currentEGB = myStringEGB.Split(',')[2];
            string resultEGB = myStringEGB.Split(',')[1];
            string resistanceEGB = myStringEGB.Split(',')[3];
            string timeEGB = myStringEGB.Split(',')[4];
            string maxEGB = ResLimit.Remove(0, 12);
            string minEGB = "0";
            string date = DateTime.UtcNow.ToString("D");
            string time = DateTime.UtcNow.ToLongTimeString();
            x.Cells[27, 2] = date;
            x.Cells[27, 4] = time;
            x.Cells[42, 1] = ClassOne;
            x.Cells[42, 4] = minEGB + "mohm";
            x.Cells[42, 5] = maxEGB + "mohm";
            x.Cells[42, 6] = currentEGB;
            x.Cells[42, 7] = resistanceEGB;
            x.Cells[42, 8] = timeEGB;
            x.Cells[42, 9] = resultEGB;
            StopTest();
            InsulationResistanceTest();             
            serGWI.WriteLine("MEAS?");
            string myStringINS = Read();
            string VoltageINS = myStringINS.Split(',')[2];
            string resultINS = myStringINS.Split(',')[1];
            string resistanceINS = myStringINS.Split(',')[3];
            string timeINS = myStringINS.Split(',')[4];
            string maxINS = "9999";
            string LowINS = "1";
            x.Cells[49, 1] = ClassOne;
            x.Cells[49, 4] = LowINS.Trim() + "Mohm";
            x.Cells[49, 5] = maxINS.Trim() + "Mohm";
            x.Cells[49, 6] = VoltageINS;
            x.Cells[49, 7] = resistanceINS.Trim();
            x.Cells[49, 8] = timeINS;
            x.Cells[49, 9] = resultINS;
            StopTest();
        }

        public void Class_2()
        {
            //opens a excel speadsheet prewriten template 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"\\guimain\ATE\CalTestGUI\MasterPatTest.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            excel.Visible = true;

            string visPass = "PASS";
            //string MbsPas = "PASS";
            string ClassTwo = "Class 2";
            x.Cells[36, 9] = visPass;
            //x.Cells[37, 9] = MbsPas;
            InsulationResistanceTest();
            //runs the main part of the the test    
            serGWI.WriteLine("MEAS?");
            string myStringINS = Read();
            string VoltageINS = myStringINS.Split(',')[2];
            string resultINS = myStringINS.Split(',')[1];
            string resistanceINS = myStringINS.Split(',')[3];
            string timeINS = myStringINS.Split(',')[4];
            string maxINS = "9999";
            string LowINS = "1";
            string date = DateTime.UtcNow.ToString("D");
            string time = DateTime.UtcNow.ToLongTimeString();
            x.Cells[27, 2] = date;
            x.Cells[27, 4] = time;
            x.Cells[49, 1] = ClassTwo;
            x.Cells[49, 4] = LowINS.Trim() + "Mohm";
            x.Cells[49, 5] = maxINS.Trim() + "Mohm";
            x.Cells[49, 6] = VoltageINS;
            x.Cells[49, 7] = resistanceINS.Trim();
            x.Cells[49, 8] = timeINS;
            x.Cells[49, 9] = resultINS;
            StopTest();
        }

        public void IecLeadTest()
        {
            //standard iec lead test
            GroundBondTest();
            serGWI.WriteLine("MEAS?");
            string ShowMeas = Read();
            MessageBox.Show(ShowMeas);
            InsulationResistanceTest();
            serGWI.WriteLine("MEAS?");
            ShowMeas = Read();
            MessageBox.Show(ShowMeas);
            StopTest();
        }
    }
}