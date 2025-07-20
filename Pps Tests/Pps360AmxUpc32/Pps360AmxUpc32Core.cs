using System;

namespace CalTestGUI.Pps360AmxUpc32
{
    public class Pps360AmxUpc32Core
    {
        private readonly Pps360AmxUpc32Communication _communication;

        public string VisaName { get; set; } = string.Empty;
        public string TestFreq { get; set; } = string.Empty;
        public string Company { get; set; } = string.Empty;
        public string CertNum { get; set; } = string.Empty;
        public string CaseNum { get; set; } = string.Empty;
        public string SerNum { get; set; } = string.Empty;
        public string UserFreq { get; set; } = string.Empty;
        public string PlantNum { get; set; } = string.Empty;

        public Pps360AmxUpc32Core(Pps360AmxUpc32Communication communication)
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
            _communication.Write_PPS(":OUTP,OFF");
            _communication.WriteN4l("KEYBOARD,ENABLE");
            _communication.Write_PPS("*GTL");
            _communication.Write_PPS("*RST");
            _communication.WriteN4l("*RST");
            _communication.ClosePort();
        }
    }
}
