using System.Windows.Forms;

namespace CalTestGUI.Pps360AsxUpc3
{
    public partial class Pps360AsxUpc3Core
    {
        protected void PpsLoopOneAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT1?");
                var result = Read_Pps();
                if (double.TryParse(result, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 10;
        }

        protected void PpsLoopTwoAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT2?");
                var result = Read_Pps();
                if (double.TryParse(result, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 10;
        }

        protected void PpsLoopThreeAcTest()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VOLT3?");
                var result = Read_Pps();
                if (double.TryParse(result, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 10;
        }

        protected void PpsLoopAcLineLine()
        {
            pps1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                Write_PPS(":MEAS:VLL1?");
                var result = Read_Pps();
                if (double.TryParse(result, out double value))
                {
                    pps1 += value;
                }
            }
            pps1 /= 10;
        }

        protected void N4lMeasOne()
        {
            meas1 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseA = Read_N4l().Split(',')[0];
                if (double.TryParse(PhaseA, out double value))
                {
                    meas1 += value;
                }
            }
            meas1 /= 10;
        }

        protected void N4lMeasTwo()
        {
            meas2 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseB = Read_N4l().Split(',')[1];
                if (double.TryParse(PhaseB, out double value))
                {
                    meas2 += value;
                }
            }
            meas2 /= 10;
        }

        protected void N4lMeasThree()
        {
            meas3 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseC = Read_N4l().Split(',')[2];
                if (double.TryParse(PhaseC, out double value))
                {
                    meas3 += value;
                }
            }
            meas3 /= 10;
        }

        protected void N4lMeasFour()
        {
            meas4 = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                Frequency = Read_N4l().Split(',')[3];
                if (double.TryParse(Frequency, out double value))
                {
                    meas4 += value;
                }
            }
            meas4 /= 10;
        }

        protected void N4lMeasLineLine()
        {
            measlL = 0;
            for (int i = 0; i < 10; i++)
            {
                Application.DoEvents();
                WriteN4l("MULTIL?");
                PhaseAb = Read_N4l().Split(',')[4];
                if (double.TryParse(PhaseAb, out double value))
                {
                    measlL += value;
                }
            }
            measlL /= 10;
        }
    }
}
