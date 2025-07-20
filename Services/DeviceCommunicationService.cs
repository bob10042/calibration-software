using System;
using System.IO.Ports;
using System.Threading;
using NationalInstruments.Visa;
using Ivi.Visa;
using System.Windows.Forms;

namespace CalTestGUI.Services
{
    public class DeviceCommunicationService : IDisposable
    {
        private readonly SerialPort serialPort;
        private MessageBasedSession? mbSession;

        public DeviceCommunicationService()
        {
            serialPort = new SerialPort
            {
                BaudRate = 9600,
                Parity = Parity.None,
                StopBits = StopBits.One,
                DataBits = 8,
                Handshake = Handshake.None,
                RtsEnable = true,
                DtrEnable = true,
                WriteTimeout = 2000,
                ReadTimeout = 2000
            };
        }

        public void OpenVisaSession(string visaAddress)
        {
            if (string.IsNullOrEmpty(visaAddress))
            {
                throw new ArgumentException("Visa address is required");
            }

            using var rmSession = new ResourceManager();
            try
            {
                mbSession = (MessageBasedSession)rmSession.Open(visaAddress);
                mbSession.TimeoutMilliseconds = 5000;
                mbSession.TerminationCharacter = 0xA;
                mbSession.SendEndEnabled = false;
                mbSession.TerminationCharacterEnabled = true;

                WritePPS("*IDN?");
                ReadPPS();
                Thread.Sleep(1000);
            }
            catch (InvalidCastException)
            {
                throw new InvalidOperationException("Resource selected must be a message-based session");
            }
        }

        public void OpenSerialPort(string portName)
        {
            serialPort.PortName = portName;
            serialPort.Open();
        }

        public void WritePPS(string command)
        {
            if (mbSession == null) return;
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        public string ReadPPS()
        {
            try
            {
                if (mbSession == null) return string.Empty;
                string messageP = InsertCommonEscapeSequences(mbSession.RawIO.ReadString());
                return messageP.Replace(@"\n", "").Replace(@"\r", "").Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }

        public void WriteN4l(string command)
        {
            serialPort.WriteLine(command);
            Thread.Sleep(200);
        }

        public string ReadN4l()
        {
            try
            {
                string messageN = serialPort.ReadLine();
                return messageN.Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
            }
        }

        private string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        private string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "\\n").Replace("\r", "\\r");
        }

        public void Dispose()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
            }
            if (mbSession != null)
            {
                mbSession.Dispose();
            }
        }

        public bool IsSerialPortOpen => serialPort.IsOpen;
        public string SerialPortName => serialPort.PortName;
    }
}
