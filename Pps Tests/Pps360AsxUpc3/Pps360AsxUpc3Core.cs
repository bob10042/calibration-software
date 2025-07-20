using NationalInstruments.Visa;
using Ivi.Visa;
using System;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Pps360AsxUpc3
{
    public partial class Pps360AsxUpc3Core
    {
        private SerialPort SerWriteN4l = new SerialPort();
        public MessageBasedSession? mbSession = null;
        
        // Configuration properties
        public string? visaName;
        public string? TestFreq;
        public string? Company;
        public string? CertNum;
        public string? CaseNum;
        public string? serNum;
        public string? UserFreq;
        public string? PlantNum;

        // Measurement values
        protected string? date = null;
        protected string? time = null;
        protected string? PhaseA = null;
        protected string? PhaseB = null;
        protected string? PhaseC = null;
        protected string? Frequency = null;
        protected string? PhaseAb = null;       

        // Measurement results
        protected double meas1 = 0;
        protected double meas2 = 0;
        protected double meas3 = 0;
        protected double meas4 = 0;
        protected double measlL = 0;
        protected double pps1 = 0;
        protected double cellvalue = 0;

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Form load initialization if needed
        }

        protected string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        protected string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "\\n").Replace("\r", "\\r");
        }

        protected void CustomRest()
        {
            Thread.Sleep(2000);
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

        public bool IsOpen_Newtowns_4th
        {
            get => SerWriteN4l.IsOpen;
        }

        public void Close()
        {
            SerWriteN4l.Close();
        }
    }
}
