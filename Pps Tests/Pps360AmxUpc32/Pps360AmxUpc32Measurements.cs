using System;
using System.Windows.Forms;

namespace CalTestGUI.Pps360AmxUpc32
{
    public class Pps360AmxUpc32Measurements
    {
        private readonly Pps360AmxUpc32Communication _communication;
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

        public Pps360AmxUpc32Measurements(Pps360AmxUpc32Communication communication)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
        }

        public double GetLastPpsMeasurement() => pps1;
        public double GetLastMeas1() => meas1;
        public double GetLastMeas2() => meas2;
        public double GetLastMeas3() => meas3;
        public double GetLastMeas4() => meas4;
        public double GetLastMeasLL() => measlL;

        public void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());                
            }
            pps1 /= 15;
        }

        public void PpsLoopTwoAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT2?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());          
            }
            pps1 /= 15;
        }

        public void PpsLoopThreeAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VOLT3?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());
            }
            pps1 /= 15;
        }

        public void PpsLoopAcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 15; i++)
            {
                Application.DoEvents();
                _communication.Write_PPS(":MEAS:VLL1?");
                pps1 += Convert.ToDouble(_communication.Read_Pps());          
            }
            pps1 /= 15;
        }

        public void N4lMeasOne()
        {
            meas1 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseA = _communication.Read_N4l().Split(',')[0];
                meas1 += Convert.ToDouble(PhaseA);    
            }
            meas1 /= 30;
        }

        public void N4lMeasTwo()
        {
            meas2 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseB = _communication.Read_N4l().Split(',')[1];
                meas2 += Convert.ToDouble(PhaseB);
            }
            meas2 /= 30;
        }

        public void N4lMeasThree()
        {
            meas3 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseC = _communication.Read_N4l().Split(',')[2];
                meas3 += Convert.ToDouble(PhaseC);
            }
            meas3 /= 30;
        }

        public void N4lMeasFour()
        {
            meas4 = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                Frequency = _communication.Read_N4l().Split(',')[3];
                meas4 += Convert.ToDouble(Frequency);
            }
            meas4 /= 30;
        }

        public void N4lMeasLineLine()
        {
            measlL = 0;
            for (int i = 0; i < 30; i++)
            {
                Application.DoEvents();
                _communication.WriteN4l("MULTIL?");
                PhaseAb = _communication.Read_N4l().Split(',')[4];
                measlL += Convert.ToDouble(PhaseAb);
            }
            measlL /= 30;
        }

        public (string KinA, string KinB, string KinC, string KexA, string KexB, string KexC, string KiA, string KiB, string KiC, string K1P) ReadKFactors()
        {
            _communication.Write_PPS(":CAL:KFACTORS:ALL?");
            string K_Factors = _communication.Read_Pps();
            string[] factors = K_Factors.Split(',');
            
            return (
                factors[0], // KinA
                factors[1], // KinB
                factors[2], // KinC
                factors[3], // KexA
                factors[4], // KexB
                factors[5], // KexC
                factors[6], // KiA
                factors[7], // KiB
                factors[8], // KiC
                factors[9]  // K1P
            );
        }
    }
}
