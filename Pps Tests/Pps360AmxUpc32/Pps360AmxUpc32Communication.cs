using NationalInstruments.Visa;
using Ivi.Visa;
using System;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;

namespace CalTestGUI.Pps360AmxUpc32
{
    public class Pps360AmxUpc32Communication
    {
        private SerialPort SerWriteN4l = new SerialPort();
        public MessageBasedSession mbSession = null;

        public void OpenSession(string visaName)
        {
            if (string.IsNullOrEmpty(visaName))
            {
                MessageBox.Show("Select Visa Address");
                return;
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

        public void Write_PPS(string command)
        {
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        public string Read_Pps()
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
        }

        public string PortName_Newtowns_4th
        {
            get { return SerWriteN4l.PortName; }
            set { SerWriteN4l.PortName = value; }
        }

        public bool IsOpen_Newtowns_4th
        {
            get { return SerWriteN4l.IsOpen; }
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

        public void WriteN4l(string command)
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

        private string ReplaceCommonEscapeSequences(string s)
        {
            return s.Replace("\\n", "\n").Replace("\\r", "\r");
        }

        private string InsertCommonEscapeSequences(string s)
        {
            return s.Replace("\n", "\\n").Replace("\r", "\\r");
        }
    }
}
