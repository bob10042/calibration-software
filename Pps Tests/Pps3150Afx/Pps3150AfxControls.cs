using System;
using System.Threading;

namespace CalTestGUI.Pps3150Afx
{
    public class Pps3150AfxControls
    {
        private readonly Pps3150AfxCommunication _communication;
        private readonly Pps3150AfxCore _core;

        public Pps3150AfxControls(Pps3150AfxCommunication communication, Pps3150AfxCore core)
        {
            _communication = communication ?? throw new ArgumentNullException(nameof(communication));
            _core = core ?? throw new ArgumentNullException(nameof(core));
        }

        public void Newtons4thControls()
        {
            _communication.WriteN4l("*RST");
            _communication.WriteN4l("TRG");                      // initiates a new measurement, resets the range and smoothing
            _communication.WriteN4l("KEYBOARD,DISABLE");         // disable keyboard
            _communication.WriteN4l("SYST;POWER");               // selects power meter
            _communication.WriteN4l("COUPLI,PHASE1,AC+DC");      // coupling ac + dc
            _communication.WriteN4l("COUPLI,PHASE2,AC+DC");      // coupling ac + dc
            _communication.WriteN4l("COUPLI,PHASE3,AC+DC");      // coupling ac + dc
            _communication.WriteN4l("APPLIC,NORMAL,SPEED3");     // select application mode
            _communication.WriteN4l("BANDWI,WIDE");              // bandwidth @ 3MHz
            _communication.WriteN4l("DATALO,DISABLE");           // datalog disabled
            _communication.WriteN4l("RESOLU,HIGH");              // sets the data resolution                             
            _communication.WriteN4l("EFFICI,0");                 // efficiency calculation disabled
            _communication.WriteN4l("FAST,ON");                  // this will disable the screen for faster resolutions 
            _communication.WriteN4l("FREQFI,OFF");               // turns the frequency filter off                     
            _communication.WriteN4l("SPEED,HIGH");               // slows the speed down 
            _communication.WriteN4l("DISPLAY,VOLTAGE");          // display voltage mode
            _communication.WriteN4l("ZOOM,1,3,5,6,7");           // sets the zoom level for the display
            Thread.Sleep(5000);
            _communication.WriteN4l("MULTIL,0");                 // reset multilogs 
            _communication.WriteN4l("MULTIL,1,1,50");            // index 1 == phase1, rms voltage
            _communication.WriteN4l("MULTIL,2,2,50");            // index 2 == phase2, rms voltage
            _communication.WriteN4l("MULTIL,3,3,50");            // index 3 == phase3, rms voltage       
            _communication.WriteN4l("MULTIL,4,1,1");             // index 4 == phase1, frequency
            _communication.WriteN4l("MULTIL,5,1,79");            // index 5 == phase-phase, rms voltage 
            _communication.WriteN4l("MULTIL,6,1,58");            // index 6 == phase1, dc voltage
            _communication.WriteN4l("MULTIL,7,2,58");            // index 7 == phase2, dc voltage
            _communication.WriteN4l("MULTIL,8,3,58");            // index 8 == phase3, dc voltage
        }

        public void ThreePhaseControlsAC()
        {
            _communication.Write_PPS("OUTPUT:AUTO,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT:MODE,AC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:FREQ,400;:VOLT:AC:LIM:MIN,0;" +
                ":VOLT:AC:LIM:MAX,300;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void ThreePhaseFreqRespControls()
        {
            _communication.Write_PPS("OUTPUT:AUTO,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT:MODE,AC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:FREQ,400;:VOLT:AC:LIM:MIN,0;" +
                ":VOLT:AC:LIM:MAX,150;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void ThreePhaseControlsDC()
        {
            _communication.Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,3;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:DC:LIM:MIN,0;:VOLT:DC:LIM:MAX,425;:COUPL,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void SplitPhaseAcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":PROG:EXEC,0;:VOLT:MODE,AC;:FORM,2;:FREQ,400;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,1500;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPL,DIRECT;:PHAS2,180;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void SplitPhaseDcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,2;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":PHAS2,180;:RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void SinglePhaseAcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":PROG:EXEC,0;:VOLT:MODE,AC;:FORM,1;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");
        }

        public void SinglePhaseDcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:INIT,OFF;:*LLO,");
            _communication.Write_PPS(":PROG:EXEC,0;:VOLT:MODE,DC;:FORM,1;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":VOLT:LIM:MIN,0;:VOLT:LIM:MAX,850;:COUPL,DIRECT;" +
                ":RANG,1;:RAMP,0;:OUTP,ON");
        }
    }
}
