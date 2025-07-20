using System;
using System.Windows.Forms;

namespace CalTestGUI.Pps3150Afx
{
    public class Pps3150AfxMeasurements
    {
        private readonly Pps3150AfxCommunication _communication;
        private double meas1 = 0;
        private double meas2 = 0;
        private double meas3 = 0;
        private double meas4 = 0;
        private double measlL = 0;
        private double pps1 = 0;

        private string PhaseA = null;
        private string PhaseB = null;
        private string PhaseC = null;
        private string Frequency = null;
        private string PhaseAb = null;

        public Pps3150AfxMeasurements(Pps3150AfxCommunication communication)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
        }

        public double GetLastPpsMeasurement() => pps1;
        public double GetLastMeas1() => meas1;
        public double GetLastMeas2() => meas2;
        public double GetLastMeas3() => meas3;
        public double GetLastMeas4() => meas4;
        public double GetLastMeasLL() => measlL;

        // AC Measurements
        public void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:AC1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        public void PpsLoopTwoAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:AC2?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        public void PpsLoopThreeAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:AC3?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        // DC Measurements
        public void PpsLoopOneDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:DC1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        public void PpsLoopTwoDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:DC2?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        public void PpsLoopThreeDcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT:DC3?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        // Line-Line Measurements
        public void PpsLoopDcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VLL1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        public void PpsLoopAcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VLL1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 10;
        }

        // N4L Measurements
        public void N4lMeasOneAc()
        {
            meas1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseA = _communication.Read_N4l().Split(',')[0];
                meas1 += Convert.ToDouble(PhaseA);
            }
            meas1 /= 10;
        }

        public void N4lMeasTwoAc()
        {
            meas2 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseB = _communication.Read_N4l().Split(',')[1];
                meas2 += Convert.ToDouble(PhaseB);
            }
            meas2 /= 10;
        }

        public void N4lMeasThreeAc()
        {
            meas3 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseC = _communication.Read_N4l().Split(',')[2];
                meas3 += Convert.ToDouble(PhaseC);
            }
            meas3 /= 10;
        }

        public void N4lMeasOneDc()
        {
            meas1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseA = _communication.Read_N4l().Split(',')[5];
                meas1 += Convert.ToDouble(PhaseA);
            }
            meas1 /= 10;
        }

        public void N4lMeasTwoDc()
        {
            meas2 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseB = _communication.Read_N4l().Split(',')[6];
                meas2 += Convert.ToDouble(PhaseB);
            }
            meas2 /= 10;
        }

        public void N4lMeasThreeDc()
        {
            meas3 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseC = _communication.Read_N4l().Split(',')[7];
                meas3 += Convert.ToDouble(PhaseC);
            }
            meas3 /= 10;
        }

        public void N4lMeasFour()
        {
            meas4 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                Frequency = _communication.Read_N4l().Split(',')[3];
                meas4 += Convert.ToDouble(Frequency);
            }
            meas4 /= 10;
        }

        public void N4lMeasLineLine()
        {
            measlL = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseAb = _communication.Read_N4l().Split(',')[4];
                measlL += Convert.ToDouble(PhaseAb);
            }
            measlL /= 10;
        }
    }
}
