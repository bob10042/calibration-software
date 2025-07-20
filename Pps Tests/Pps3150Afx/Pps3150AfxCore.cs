using System;

namespace CalTestGUI.Pps3150Afx
{
    public class Pps3150AfxCore
    {
        private readonly Pps3150AfxCommunication _communication;

        public string VisaName { get; set; } = string.Empty;
        public string TestFreq { get; set; } = string.Empty;
        public string Company { get; set; } = string.Empty;
        public string CertNum { get; set; } = string.Empty;
        public string CaseNum { get; set; } = string.Empty;
        public string SerNum { get; set; } = string.Empty;
        public string UserFreq { get; set; } = string.Empty;
        public string PlantNum { get; set; } = string.Empty;

        public Pps3150AfxCore(Pps3150AfxCommunication communication)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
        }

        public void Initialize()
        {
            if (string.IsNullOrEmpty(VisaName))
            {
                throw new InvalidOperationException("VisaName must be set before initialization");
            }

            _communication.OpenSession(VisaName);
        }

        public void CustomRest()
        {
            System.Threading.Thread.Sleep(2000);
        }

        public void Cleanup()
        {
            try
            {
                _communication.Write_PPS(":OUTP,OFF");
                _communication.WriteN4l("KEYBOARD,ENABLE");
                _communication.Write_PPS("*GTL");
                _communication.Write_PPS("*RST");
                _communication.WriteN4l("*RST");
            }
            finally
            {
                _communication.ClosePort();
            }
        }

        public void SetVoltageAC(double voltage)
        {
            _communication.Write_PPS($":VOLT:AC,{voltage}");
            System.Threading.Thread.Sleep(200);
        }

        public void SetVoltageDC(double voltage)
        {
            _communication.Write_PPS($":VOLT:DC,{voltage}");
            System.Threading.Thread.Sleep(200);
        }

        public void SetFrequency(double frequency)
        {
            _communication.Write_PPS($":FREQ,{frequency}");
            System.Threading.Thread.Sleep(200);
        }
    }
}
