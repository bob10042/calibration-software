using System;
using NationalInstruments.Visa;
using System.Threading;

namespace CalTestGUI
{
    public abstract class BaseMeter
    {
        protected MessageBasedSession? mbSession;
        private ResourceManager? rmSession;
        public string VisaAddress { get; set; } = string.Empty;
        public string SerialNumber { get; set; } = string.Empty;
        public string CalibrationCertNumber { get; set; } = string.Empty;

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
            if (mbSession == null)
                throw new InvalidOperationException("Session not initialized. Call OpenSession first.");
                
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        protected string Read()
        {
            if (mbSession == null)
                throw new InvalidOperationException("Session not initialized. Call OpenSession first.");
                
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
                if (string.IsNullOrEmpty(VisaAddress))
                    throw new InvalidOperationException("VISA address not set");

                CloseSession(); // Ensure any existing session is properly closed
                
                rmSession = new ResourceManager();
                mbSession = (MessageBasedSession)rmSession.Open(VisaAddress);
                mbSession.TimeoutMilliseconds = 5000;
                mbSession.TerminationCharacter = 0xA;
                mbSession.SendEndEnabled = false;
                mbSession.TerminationCharacterEnabled = true;
            }
            catch (Exception ex)
            {
                CloseSession(); // Clean up on failure
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
            
            if (rmSession != null)
            {
                rmSession.Dispose();
                rmSession = null;
            }
        }

        public abstract string GetIdentification();
        public abstract string Measure(string parameter);
    }
}
