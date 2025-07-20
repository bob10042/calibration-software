using System;
using System.Threading;
using System.Windows.Forms;
using System.IO.Ports;
using Ivi.Visa;
using NationalInstruments.Visa;

namespace CalTestGUI.Pps360AsxUpc3
{
    public partial class Pps360AsxUpc3Core
    {
        public void OpenSession()
        {
            if (visaName == null)
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

        protected void Write_PPS(string command)
        {
            if (mbSession == null) return;
            mbSession.RawIO.Write(ReplaceCommonEscapeSequences(command + "\n"));
            Thread.Sleep(200);
        }

        protected string Read_Pps()
        {
            try
            {
                if (mbSession == null) return string.Empty;
                string messageP = InsertCommonEscapeSequences((mbSession.RawIO.ReadString()));
                return messageP.Replace(@"\n", "").Replace(@"\r", "").Trim();
            }
            catch (TimeoutException)
            {
                return "Timeout";
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

        protected void WriteN4l(string command)
        {
            SerWriteN4l.WriteLine(command);
            Thread.Sleep(200);
        }

        protected string Read_N4l()
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
    }
}
