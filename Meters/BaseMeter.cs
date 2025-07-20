using System;
using NationalInstruments.Visa;
using System.Threading;

namespace CalTestGUI
{
    public abstract class BaseMeter
    {
        protected MessageBasedSession mbSession;
        public string VisaAddress { get; set; }
        public string SerialNumber { get; set; }
        public string CalibrationCertNumber { get; set; }

        protected string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        protected string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "\\n").Replace("\r", "\\r");
        }

        protected void Write(string command)
        {
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        protected string Read()
        {
            try
            {
                string message = InsertCommonEscapeSequences(mbSession.RawIO.ReadString());
                return message.Replace(@"\n", "").Replace(@"\r", "").Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }

        public virtual void OpenSession()
        {
            try
            {
                using (var rmSession = new ResourceManager())
                {
                    mbSession = (MessageBasedSession)rmSession.Open(VisaAddress);
                    mbSession.TimeoutMilliseconds = 5000;
                    mbSession.TerminationCharacter = 0xA;
                    mbSession.SendEndEnabled = false;
                    mbSession.TerminationCharacterEnabled = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to open session to meter: {ex.Message}");
            }
        }

        public virtual void CloseSession()
        {
            if (mbSession != null)
            {
                mbSession.Dispose();
                mbSession = null;
            }
        }

        public abstract string GetIdentification();
        public abstract string Measure(string parameter);
    }
}
