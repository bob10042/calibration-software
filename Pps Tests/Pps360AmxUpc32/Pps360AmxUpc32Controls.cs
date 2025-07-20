using System;

namespace CalTestGUI.Pps360AmxUpc32
{
    public class Pps360AmxUpc32Controls
    {
        private readonly Pps360AmxUpc32Communication _communication;
        private readonly Pps360AmxUpc32Core _core;

        public Pps360AmxUpc32Controls(Pps360AmxUpc32Communication communication, Pps360AmxUpc32Core core)
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
            System.Threading.Thread.Sleep(5000);     
            _communication.WriteN4l("MULTIL,0");                 // reset multilogs 
            _communication.WriteN4l("MULTIL,1,1,50");            // index 1 == phase1, rms voltage
            _communication.WriteN4l("MULTIL,2,2,50");            // index 2 == phase2, rms voltage
            _communication.WriteN4l("MULTIL,3,3,50");            // index 3 == phase3, rms voltage       
            _communication.WriteN4l("MULTIL,4,1,1");             // index 4 == phase1, frequency
            _communication.WriteN4l("MULTIL,5,1,79");            // index 5 == phase-phase, rms voltage 
            _communication.WriteN4l("REZERO");                   // request the dsp to re-compensate for dc offset
        }

        public void ThreePhaseControls()
        {                                       
            _communication.Write_PPS("*LLO");                            
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,3;:FREQ," + _core.UserFreq.Trim() +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHAS2,120;" +
                ":PHAS3,240;:RANG,0;:RAMP,0;:OUTP,ON");
            _core.CustomRest();
        }

        public void SpitPhaseAcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,2;:FREQ,50;" +
                ":VOLT:ALC,OFF;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            _core.CustomRest();
        }

        public void SinglePhaseAcControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:VOLT,0;:FORM,1;:FREQ," + _core.UserFreq.Trim() +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;" +
                ":RANG,0;:RAMP,0;:OUTP,ON");
            _core.CustomRest();
        }

        public void ThreePhaseFreqRespControls()
        {
            _communication.Write_PPS(":OUTP,OFF;:PROG:EXEC,0;:FORM,3;:VOLT,0;:FREQ,50;" +
                ":VOLT:ALC,ON;:CURR:LIM,10;:WAVEFORM,1;:SENS:PATH,0;" +
                ":FREQ:LIM:MIN,45;:FREQ:LIM:MAX,5000;:VOLT:LIM:MIN,0;" +
                ":VOLT:LIM:MAX,600;:COUPLING,DIRECT;:PHASe2,120;" +
                ":PHASe3,240;:RANG,0;:RAMP,0;:VOLT,120;:OUTP,ON");
            _core.CustomRest();
        }
    }
}
